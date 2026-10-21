import email
from datetime import date, timedelta

import pytest

import email_sender
import excel_manager
import scheduler
from email_sender import parse_email_draft, send_followup_email
from scheduler import compute_followup_date


# ── Pure helpers ────────────────────────────────────────────────────────

class TestParseEmailDraft:
    def test_splits_subject_and_body(self):
        subject, body = parse_email_draft("Subject: Following up\n\nHi Jane,\nThanks.")
        assert subject == "Following up"
        assert body == "Hi Jane,\nThanks."

    def test_default_subject_when_missing(self):
        subject, body = parse_email_draft("Hi Jane,\nThanks.")
        assert subject == "Following up on my application"
        assert body.startswith("Hi Jane")

    def test_only_first_subject_line_is_used(self):
        subject, body = parse_email_draft("Subject: A\nSubject: B\nBody")
        assert subject == "A"
        assert "Subject: B" in body


class TestComputeFollowupDate:
    @pytest.mark.parametrize("count,expected", [(0, "2026-10-08"), (1, "2026-10-15"), (2, "2026-10-22"), (9, "2026-10-22")])
    def test_increasing_intervals_capped_at_third(self, count, expected):
        assert compute_followup_date("2026-10-01", count) == expected

    @pytest.mark.parametrize("bad", ["", "01.10.2026", "not a date"])
    def test_unparseable_date_returns_empty(self, bad):
        assert compute_followup_date(bad, 0) == ""


class TestHeaderMapping:
    @pytest.mark.parametrize("raw,expected", [
        ("Company Name", "Company"),
        ("Date of applied", "Applied Date"),
        ("Repsonse received", "Status"),
        ("  Recruiter Email ", "Recruiter Email"),
        ("Unnamed: 0", "ID"),
        ("Something Else", "Something Else"),
    ])
    def test_aliases(self, raw, expected):
        assert excel_manager._map_header_name(raw) == expected


# ── Excel round trip ────────────────────────────────────────────────────

@pytest.fixture
def workbook(tmp_path):
    path = tmp_path / "apps.xlsx"
    excel_manager.create_default_excel(str(path))
    return path


class TestExcelManager:
    def test_default_workbook_has_one_sample_row(self, workbook):
        apps = excel_manager.load_applications(str(workbook))
        assert len(apps) == 1
        assert apps[0]["Company"] == "Example Corp"

    def test_add_then_update_application(self, workbook):
        new_id = excel_manager.add_application(str(workbook), {"Company": "Acme", "Role": "Dev", "Status": "Applied"})
        assert new_id == 2

        assert excel_manager.save_application_update(str(workbook), "2", {"Status": "Rejected"}) is True
        acme = next(a for a in excel_manager.load_applications(str(workbook)) if a["Company"] == "Acme")
        assert acme["Status"] == "Rejected"
        assert acme["Email Sent"] == "No"

    def test_update_unknown_id_reports_failure(self, workbook):
        assert excel_manager.save_application_update(str(workbook), "999", {"Status": "Rejected"}) is False


# ── Email sending (SMTP faked) ──────────────────────────────────────────

class FakeSMTP:
    sent = []

    def __init__(self, host, port, timeout=None):
        self.host, self.port, self.timeout = host, port, timeout

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False

    def ehlo(self):
        pass

    def starttls(self):
        pass

    def login(self, user, password):
        pass

    def sendmail(self, sender, recipients, message):
        FakeSMTP.sent.append((sender, recipients, message))


@pytest.fixture
def smtp(monkeypatch):
    FakeSMTP.sent = []
    monkeypatch.setattr(email_sender.smtplib, "SMTP", FakeSMTP)
    monkeypatch.setenv("EMAIL_SENDER", "me@example.com")
    monkeypatch.setenv("EMAIL_PASSWORD", "test-only")
    return FakeSMTP


class TestSendFollowupEmail:
    def test_requires_configuration(self, monkeypatch):
        monkeypatch.delenv("EMAIL_SENDER", raising=False)
        ok, message = send_followup_email("a@example.com", "s", "b")
        assert not ok and "not configured" in message

    @pytest.mark.parametrize("bad", ["", "no-at-sign", "a@example.com\r\nBcc: x@evil.test", "a@x.test, b@y.test"])
    def test_rejects_invalid_recipients(self, smtp, bad):
        ok, _ = send_followup_email(bad, "s", "b")
        assert not ok
        assert smtp.sent == []

    def test_sends_to_recipient_and_valid_cc(self, smtp):
        ok, _ = send_followup_email("rec@example.com", "Hello", "Body", cc_email="hr@example.com")
        assert ok
        _, recipients, _ = smtp.sent[0]
        assert recipients == ["rec@example.com", "hr@example.com"]

    def test_invalid_cc_is_ignored(self, smtp):
        ok, _ = send_followup_email("rec@example.com", "Hello", "Body", cc_email="not-an-email")
        assert ok
        assert smtp.sent[0][1] == ["rec@example.com"]

    def test_subject_newlines_cannot_inject_headers(self, smtp):
        send_followup_email("rec@example.com", "Hi\r\nBcc: spy@evil.test", "Body")
        raw = smtp.sent[0][2]
        assert "\nBcc:" not in raw

    def test_html_part_escapes_body(self, smtp):
        send_followup_email("rec@example.com", "Hi", "<script>alert(1)</script>")
        message = email.message_from_string(smtp.sent[0][2])
        parts = {p.get_content_type(): p.get_payload(decode=True).decode() for p in message.walk()
                 if not p.is_multipart()}
        assert "<script>" in parts["text/plain"]          # plain text is sent as written
        assert "<script>" not in parts["text/html"]       # but never as live markup
        assert "&lt;script&gt;" in parts["text/html"]


# ── Scheduler (Claude and SMTP faked) ───────────────────────────────────

def _app(**overrides):
    base = {"ID": "1", "Company": "Acme", "Role": "Dev", "Status": "Applied",
            "Applied Date": (date.today() - timedelta(days=10)).isoformat(),
            "Follow Up Date": (date.today() - timedelta(days=1)).isoformat(),
            "Follow Up Count": "0", "Recruiter Email": "rec@example.com", "HR Email": ""}
    base.update(overrides)
    return base


class TestRunFollowupCheck:
    @pytest.fixture(autouse=True)
    def fakes(self, monkeypatch):
        self.updates = []
        self.sent = []
        monkeypatch.setattr(scheduler, "generate_followup_email", lambda **kw: f"Subject: Re {kw['role']}\n\nHi")
        monkeypatch.setattr(scheduler, "save_application_update", lambda path, i, u: self.updates.append((i, u)) or True)
        monkeypatch.setattr(scheduler, "send_followup_email", lambda **kw: (self.sent.append(kw) or True, "sent"))

    def run(self, monkeypatch, apps, dry_run=False):
        monkeypatch.setattr(scheduler, "load_applications", lambda path: apps)
        return scheduler.run_followup_check("unused.xlsx", "Alex", dry_run=dry_run)

    def test_dry_run_drafts_but_never_sends_or_writes(self, monkeypatch):
        results = self.run(monkeypatch, [_app()], dry_run=True)
        assert results[0]["status"] == "dry_run"
        assert results[0]["subject"] == "Re Dev"
        assert self.sent == [] and self.updates == []

    def test_due_application_is_sent_and_rescheduled(self, monkeypatch):
        results = self.run(monkeypatch, [_app()])
        assert results[0]["status"] == "sent"
        _, update = self.updates[-1]
        assert update["Follow Up Count"] == 1
        assert update["Email Sent"] == "Yes"

    @pytest.mark.parametrize("status", ["Rejected", "Withdrawn", "Offer Received", "Interview Scheduled"])
    def test_closed_statuses_are_skipped(self, monkeypatch, status):
        assert self.run(monkeypatch, [_app(Status=status)]) == []

    def test_future_follow_up_is_skipped(self, monkeypatch):
        future = (date.today() + timedelta(days=3)).isoformat()
        assert self.run(monkeypatch, [_app(**{"Follow Up Date": future})]) == []

    def test_missing_recruiter_email_is_reported(self, monkeypatch):
        results = self.run(monkeypatch, [_app(**{"Recruiter Email": ""})])
        assert results[0]["status"] == "skipped"
        assert self.sent == []
