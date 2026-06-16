"""
Scheduler - Checks follow-up dates and triggers emails automatically.
Runs in a background thread; checks every hour.
"""

import os
import threading
import time
from datetime import date, datetime, timedelta

import schedule

from excel_manager import load_applications, save_application_update
from ai_email_generator import generate_followup_email
from email_sender import send_followup_email, parse_email_draft


def compute_followup_date(applied_date_str: str, follow_up_count: int) -> str:
    """Calculate next follow-up date with increasing intervals."""
    try:
        applied = datetime.strptime(applied_date_str.strip(), "%Y-%m-%d").date()
    except ValueError:
        return ""

    intervals = [7, 14, 21]  # days after applied date for 1st, 2nd, 3rd follow-ups
    idx = min(follow_up_count, len(intervals) - 1)
    next_date = applied + timedelta(days=intervals[idx])
    return next_date.strftime("%Y-%m-%d")


def run_followup_check(excel_path: str, sender_name: str = "", dry_run: bool = False) -> list[dict]:
    """Check all applications and send due follow-ups. Returns list of results."""
    today = date.today().strftime("%Y-%m-%d")
    apps = load_applications(excel_path)
    results = []

    skip_statuses = {"Rejected", "Withdrawn", "Offer Received", "Interview Scheduled", "Interview Done"}

    for app in apps:
        app_id = app.get("ID", "")
        company = app.get("Company", "")
        status = app.get("Status", "Applied")
        follow_up_date = app.get("Follow Up Date", "").strip()
        follow_up_count = int(app.get("Follow Up Count", 0) or 0)
        applied_date = app.get("Applied Date", "").strip()
        recruiter_email = app.get("Recruiter Email", "").strip()

        if status in skip_statuses:
            continue

        # Auto-set follow_up_date if missing
        if not follow_up_date and applied_date:
            follow_up_date = compute_followup_date(applied_date, follow_up_count)
            save_application_update(excel_path, app_id, {"Follow Up Date": follow_up_date})

        if not follow_up_date or follow_up_date > today:
            continue

        if not recruiter_email or "@" not in recruiter_email:
            results.append({
                "id": app_id, "company": company,
                "status": "skipped", "reason": "No recruiter email"
            })
            continue

        # Generate AI email
        try:
            draft = generate_followup_email(
                company=company,
                role=app.get("Role", ""),
                recruiter_name=app.get("Recruiter Name", ""),
                applied_date=applied_date,
                follow_up_count=follow_up_count,
                notes=app.get("Notes", ""),
                sender_name=sender_name,
            )
        except Exception as e:
            results.append({
                "id": app_id, "company": company,
                "status": "error", "reason": f"AI draft failed: {e}"
            })
            continue

        subject, body = parse_email_draft(draft)

        if dry_run:
            results.append({
                "id": app_id, "company": company, "status": "dry_run",
                "to": recruiter_email, "subject": subject, "draft": draft
            })
            continue

        success, message = send_followup_email(
            to_email=recruiter_email,
            subject=subject,
            body=body,
            cc_email=app.get("HR Email", ""),
            sender_name=sender_name,
        )

        new_count = follow_up_count + 1
        next_followup = compute_followup_date(applied_date, new_count)

        updates = {
            "Follow Up Count": new_count,
            "Last Follow Up Date": today,
            "Follow Up Date": next_followup,
            "Status": "Follow Up Sent",
            "AI Draft": draft,
            "Email Sent": "Yes" if success else "Failed",
        }
        save_application_update(excel_path, app_id, updates)

        results.append({
            "id": app_id, "company": company,
            "status": "sent" if success else "failed",
            "reason": message, "to": recruiter_email,
            "subject": subject,
        })
        print(f"[{datetime.now().strftime('%H:%M')}] {company}: {message}")

    return results


class FollowUpScheduler:
    def __init__(self, excel_path: str, sender_name: str = "", check_interval_hours: int = 1):
        self.excel_path = excel_path
        self.sender_name = sender_name
        self.check_interval_hours = check_interval_hours
        self._thread = None
        self._running = False

    def start(self):
        schedule.every(self.check_interval_hours).hours.do(self._job)
        schedule.run_all()  # Run once immediately on startup

        self._running = True
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()
        print(f"Scheduler started — checking every {self.check_interval_hours} hour(s)")

    def stop(self):
        self._running = False
        schedule.clear()

    def trigger_now(self) -> list[dict]:
        return run_followup_check(self.excel_path, self.sender_name)

    def _job(self):
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M')}] Running scheduled follow-up check...")
        run_followup_check(self.excel_path, self.sender_name)

    def _loop(self):
        while self._running:
            schedule.run_pending()
            time.sleep(60)
