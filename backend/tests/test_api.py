import pytest

import app as app_module
import excel_manager


@pytest.fixture
def client(tmp_path, monkeypatch):
    path = tmp_path / "apps.xlsx"
    excel_manager.create_default_excel(str(path))
    monkeypatch.setattr(app_module, "EXCEL_PATH", str(path))
    app_module.app.config["TESTING"] = True
    return app_module.app.test_client()


def test_health(client):
    assert client.get("/api/health").json["status"] == "ok"


def test_lists_applications_with_stats(client):
    body = client.get("/api/applications").json
    assert body["stats"]["total"] == 1
    assert body["applications"][0]["Company"] == "Example Corp"


def test_unknown_origin_gets_no_cors_header(client):
    resp = client.get("/api/applications", headers={"Origin": "https://evil.example"})
    assert "Access-Control-Allow-Origin" not in resp.headers


def test_local_frontend_origin_is_allowed(client):
    resp = client.get("/api/applications", headers={"Origin": "http://localhost:5500"})
    assert resp.headers.get("Access-Control-Allow-Origin") == "http://localhost:5500"


@pytest.mark.parametrize("path", ["/api/run-scheduler", "/api/dry-run", "/api/generate-all-drafts",
                                  "/api/applications/1/send-email", "/api/applications"])
def test_writes_require_json(client, path):
    # A plain form POST is what a malicious page could send without a CORS preflight.
    resp = client.post(path, data="x=1", content_type="application/x-www-form-urlencoded")
    assert resp.status_code == 415


def test_create_and_update_application(client):
    created = client.post("/api/applications", json={"Company": "Acme", "Role": "Dev"}).json
    assert created["success"] and created["id"] == 2

    assert client.put("/api/applications/2", json={"Status": "Rejected"}).status_code == 200
    assert client.put("/api/applications/999", json={"Status": "Rejected"}).status_code == 404


def test_send_without_draft_is_rejected(client):
    resp = client.post("/api/applications/1/send-email", json={})
    assert resp.status_code == 400


def test_send_uses_given_draft_and_marks_row(client, monkeypatch):
    calls = []
    monkeypatch.setattr(app_module, "send_followup_email",
                        lambda **kw: (calls.append(kw) or True, "sent"))
    resp = client.post("/api/applications/1/send-email",
                       json={"draft": "Subject: Checking in\n\nHi Jane"})
    assert resp.status_code == 200
    assert calls[0]["subject"] == "Checking in"
    assert calls[0]["to_email"] == "jane@example.com"

    row = client.get("/api/applications").json["applications"][0]
    assert row["Email Sent"] == "Yes"
    assert row["Follow Up Count"] == "1"


def test_failed_send_returns_error_status(client, monkeypatch):
    monkeypatch.setattr(app_module, "send_followup_email", lambda **kw: (False, "SMTP down"))
    resp = client.post("/api/applications/1/send-email", json={"draft": "Hi"})
    assert resp.status_code == 502


def test_config_never_exposes_secrets(client, monkeypatch):
    monkeypatch.setenv("ANTHROPIC_API_KEY", "test-key-value")
    monkeypatch.setenv("EMAIL_PASSWORD", "test-password-value")
    text = client.get("/api/config").get_data(as_text=True)
    assert "test-key-value" not in text
    assert "test-password-value" not in text
