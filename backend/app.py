"""
Flask API backend for the Auto Follow-Up Tool (experimental prototype).

This API is separate from the WinForms desktop app in ../WinFormsApp and has
no bundled frontend. It is meant to be called from localhost only.
"""

import os
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

from dotenv import load_dotenv
from flask import Flask, jsonify, request
from flask_cors import CORS

# Load env from project root
load_dotenv(Path(__file__).parent.parent / ".env")

from excel_manager import (
    load_applications, save_application_update,
    add_application, create_default_excel, ExcelWatcher, COLUMNS, STATUS_OPTIONS,
    sync_source_status
)
from ai_email_generator import generate_followup_email, generate_bulk_drafts, get_email_prompt_template
from email_sender import send_followup_email, parse_email_draft
from scheduler import FollowUpScheduler, run_followup_check, compute_followup_date

app = Flask(__name__)

# Only local pages may call this API from a browser. With the previous
# CORS(app) any website open in the user's browser could read the
# application list (names, recruiter emails) from http://localhost:5000.
ALLOWED_ORIGINS = [
    o.strip() for o in os.getenv(
        "CORS_ALLOWED_ORIGINS", "http://localhost:5500,http://127.0.0.1:5500"
    ).split(",") if o.strip()
]
CORS(app, origins=ALLOWED_ORIGINS)

# Automatic sending is opt-in: without it the hourly scheduler never runs and
# every email goes out only through an explicit request.
AUTO_SEND_ENABLED = os.getenv("AUTO_SEND_ENABLED", "false").strip().lower() in {"1", "true", "yes"}

EXCEL_PATH = os.getenv("EXCEL_FILE_PATH", "data/job_applications.xlsx")
if not Path(EXCEL_PATH).is_absolute():
    EXCEL_PATH = str(Path(__file__).parent.parent / EXCEL_PATH)

SOURCE_EXCEL_PATH = os.getenv("EXCEL_SOURCE_PATH", "")
if SOURCE_EXCEL_PATH and not Path(SOURCE_EXCEL_PATH).is_absolute():
    SOURCE_EXCEL_PATH = str(Path(__file__).parent.parent / SOURCE_EXCEL_PATH)

SENDER_NAME = os.getenv("SENDER_NAME", "")

# In-memory cache
_applications_cache: list[dict] = []

def reload_cache():
    global _applications_cache
    _applications_cache = load_applications(EXCEL_PATH)

def get_stats(apps: list[dict]) -> dict:
    today = date.today().strftime("%Y-%m-%d")
    due_today = sum(
        1 for a in apps
        if a.get("Follow Up Date", "") <= today
        and a.get("Follow Up Date", "")
        and a.get("Status", "") not in {"Rejected", "Withdrawn", "Offer Received"}
    )
    by_status = {}
    for a in apps:
        s = a.get("Status", "Unknown")
        by_status[s] = by_status.get(s, 0) + 1

    return {
        "total": len(apps),
        "due_today": due_today,
        "by_status": by_status,
        "active": sum(1 for a in apps if a.get("Status") not in {"Rejected", "Withdrawn"}),
    }


@app.before_request
def require_json_for_writes():
    """
    State-changing requests must declare a JSON body. Browsers cannot send
    application/json cross-origin without a CORS preflight, which the CORS
    policy above rejects, so a malicious page cannot trigger sends with a
    simple form POST to localhost.
    """
    if request.method in {"POST", "PUT", "PATCH", "DELETE"} and not request.is_json:
        return jsonify({"error": "Content-Type must be application/json"}), 415
    return None


def _json_body() -> dict:
    data = request.get_json(silent=True)
    return data if isinstance(data, dict) else {}


# ──────────────────────────────────────────────
# Routes
# ──────────────────────────────────────────────

@app.get("/api/health")
def health():
    return jsonify({"status": "ok", "excel_path": EXCEL_PATH})


@app.get("/api/applications")
def get_applications():
    reload_cache()
    # Derive columns from loaded cache if possible so frontend matches Excel exactly
    cols = COLUMNS
    if _applications_cache and isinstance(_applications_cache, list) and len(_applications_cache) > 0:
        cols = list(_applications_cache[0].keys())

    return jsonify({
        "applications": _applications_cache,
        "stats": get_stats(_applications_cache),
        "columns": cols,
        "status_options": STATUS_OPTIONS,
    })


@app.post("/api/sync-source")
def sync_source():
    if not SOURCE_EXCEL_PATH or not Path(SOURCE_EXCEL_PATH).exists():
        return jsonify({"success": False, "message": "Source Excel file not configured or missing."}), 404

    updated = sync_source_status(SOURCE_EXCEL_PATH, EXCEL_PATH)
    reload_cache()
    return jsonify({"success": True, "updated": updated})


@app.post("/api/applications")
def create_application():
    data = _json_body()
    data.setdefault("Applied Date", date.today().strftime("%Y-%m-%d"))
    data.setdefault("Status", "Applied")

    # Auto-set follow-up date
    if not data.get("Follow Up Date") and data.get("Applied Date"):
        data["Follow Up Date"] = compute_followup_date(data["Applied Date"], 0)

    new_id = add_application(EXCEL_PATH, data)
    reload_cache()
    return jsonify({"success": True, "id": new_id, "message": f"Application #{new_id} added"})


@app.put("/api/applications/<app_id>")
def update_application(app_id: str):
    updates = _json_body()
    success = save_application_update(EXCEL_PATH, app_id, updates)
    if success:
        reload_cache()
        return jsonify({"success": True})
    return jsonify({"success": False, "error": "Application not found"}), 404


@app.get("/api/applications/<app_id>/generate-draft")
def generate_draft(app_id: str):
    reload_cache()
    app_data = next((a for a in _applications_cache if str(a.get("ID")) == app_id), None)
    if not app_data:
        return jsonify({"error": "Application not found"}), 404

    try:
        draft = generate_followup_email(
            company=app_data.get("Company", ""),
            role=app_data.get("Role", ""),
            recruiter_name=app_data.get("Recruiter Name", ""),
            applied_date=app_data.get("Applied Date", ""),
            follow_up_count=int(app_data.get("Follow Up Count", 0) or 0),
            notes=app_data.get("Notes", ""),
            sender_name=SENDER_NAME,
        )
        save_application_update(EXCEL_PATH, app_id, {"AI Draft": draft})
        reload_cache()
        return jsonify({"success": True, "draft": draft})
    except Exception as e:  # surfaced to the local user; the API key is never part of the message
        app.logger.exception("Draft generation failed for application %s", app_id)
        return jsonify({"error": f"Draft generation failed: {e}"}), 502


@app.post("/api/applications/<app_id>/send-email")
def send_email(app_id: str):
    reload_cache()
    app_data = next((a for a in _applications_cache if str(a.get("ID")) == app_id), None)
    if not app_data:
        return jsonify({"error": "Application not found"}), 404

    body_data = _json_body()
    draft = body_data.get("draft") or app_data.get("AI Draft", "")

    if not draft:
        return jsonify({"error": "No email draft available — generate one first"}), 400

    subject, body = parse_email_draft(draft)
    to_email = app_data.get("Recruiter Email", "")
    cc_email = app_data.get("HR Email", "")

    success, message = send_followup_email(
        to_email=to_email,
        subject=subject,
        body=body,
        cc_email=cc_email,
        sender_name=SENDER_NAME,
    )

    if success:
        today = date.today().strftime("%Y-%m-%d")
        applied_date = app_data.get("Applied Date", "")
        new_count = int(app_data.get("Follow Up Count", 0) or 0) + 1
        next_followup = compute_followup_date(applied_date, new_count)
        save_application_update(EXCEL_PATH, app_id, {
            "Email Sent": "Yes",
            "Status": "Follow Up Sent",
            "Follow Up Count": new_count,
            "Last Follow Up Date": today,
            "Follow Up Date": next_followup,
            "AI Draft": draft,
        })
        reload_cache()

    return jsonify({"success": success, "message": message}), (200 if success else 502)


@app.post("/api/run-scheduler")
def run_scheduler_now():
    results = run_followup_check(EXCEL_PATH, SENDER_NAME)
    reload_cache()
    return jsonify({"results": results, "count": len(results)})


@app.post("/api/dry-run")
def dry_run():
    from scheduler import run_followup_check
    results = run_followup_check(EXCEL_PATH, SENDER_NAME, dry_run=True)
    return jsonify({"results": results, "count": len(results)})


@app.post("/api/generate-all-drafts")
def generate_all_drafts():
    reload_cache()
    drafts = generate_bulk_drafts(_applications_cache, sender_name=SENDER_NAME)
    for app_id, draft in drafts.items():
        save_application_update(EXCEL_PATH, str(app_id), {"AI Draft": draft})
    reload_cache()
    return jsonify({"generated": len(drafts), "ids": list(drafts.keys())})


@app.get("/api/stats")
def get_stats_endpoint():
    reload_cache()
    return jsonify(get_stats(_applications_cache))


@app.get("/api/config")
def get_config():
    # include last-modified time of the excel file (if available)
    excel_mtime = ""
    try:
        if Path(EXCEL_PATH).exists():
            m = Path(EXCEL_PATH).stat().st_mtime
            excel_mtime = datetime.fromtimestamp(m).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        excel_mtime = ""

    return jsonify({
        "excel_path": EXCEL_PATH,
        "excel_modified": excel_mtime,
        "sender_name": SENDER_NAME,
        "sender_email": os.getenv("EMAIL_SENDER", ""),
        "smtp_host": os.getenv("EMAIL_SMTP_HOST", "smtp.gmail.com"),
        "follow_up_interval_days": os.getenv("FOLLOW_UP_INTERVAL_DAYS", "7"),
        "anthropic_key_set": bool(os.getenv("ANTHROPIC_API_KEY")),
        "email_configured": bool(os.getenv("EMAIL_SENDER") and os.getenv("EMAIL_PASSWORD")),
        "auto_send_enabled": AUTO_SEND_ENABLED,
        "email_template": get_email_prompt_template(),
    })


if __name__ == "__main__":
    if not Path(EXCEL_PATH).exists():
        create_default_excel(EXCEL_PATH)

    if SOURCE_EXCEL_PATH and Path(SOURCE_EXCEL_PATH).exists():
        synced = sync_source_status(SOURCE_EXCEL_PATH, EXCEL_PATH)
        if synced:
            print(f"Synced {synced} Status values from source Excel: {SOURCE_EXCEL_PATH}")

    reload_cache()

    # Start Excel file watcher
    watcher = ExcelWatcher(EXCEL_PATH, reload_cache)
    watcher.start()

    # Start automated scheduler (sends real emails without review - opt-in only)
    if not AUTO_SEND_ENABLED:
        print("Automatic sending disabled. Set AUTO_SEND_ENABLED=true in .env to enable the hourly scheduler.")
    elif os.getenv("EMAIL_SENDER") and os.getenv("EMAIL_PASSWORD"):
        scheduler = FollowUpScheduler(EXCEL_PATH, SENDER_NAME)
        scheduler.start()
    else:
        print("Email not configured — scheduler disabled. Set EMAIL_SENDER and EMAIL_PASSWORD in .env")

    port = int(os.getenv("FLASK_PORT", 5000))
    print(f"\nAuto Follow-Up Tool API running at http://127.0.0.1:{port}/api/health")
    print(f"Excel file: {EXCEL_PATH}\n")
    app.run(host="127.0.0.1", debug=False, port=port, use_reloader=False)
