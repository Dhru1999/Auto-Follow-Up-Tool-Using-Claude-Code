"""
AI Email Generator - Uses Claude to craft personalized follow-up emails.
"""

import os

import anthropic

DEFAULT_MODEL = "claude-sonnet-4-6"
MAX_NOTES_LENGTH = 500


def generate_followup_email(
    company: str,
    role: str,
    recruiter_name: str,
    applied_date: str,
    follow_up_count: int,
    notes: str = "",
    sender_name: str = "Job Applicant",
) -> str:
    # Reads ANTHROPIC_API_KEY from the environment; retries transient errors.
    client = anthropic.Anthropic(timeout=60.0, max_retries=2)
    notes = (notes or "").strip()[:MAX_NOTES_LENGTH]

    follow_up_number = follow_up_count + 1
    ordinal = {1: "first", 2: "second", 3: "third"}.get(follow_up_number, f"{follow_up_number}th")

    prompt = f"""Write a professional, concise, and warm {ordinal} follow-up email for a job application.

Details:
- Company: {company}
- Role: {role}
- Recruiter/Contact: {recruiter_name or "Hiring Team"}
- Application Date: {applied_date}
- Sender Name: {sender_name}
- Additional Context: {notes or "None"}

Requirements:
- Subject line on first line starting with "Subject: "
- 3-4 short paragraphs max
- Professional but warm tone
- Express genuine interest without being pushy
- Mention the specific role and company name
- End with a clear call to action
- Include a polite sign-off with the sender's name
- Do NOT use placeholder brackets like [Your Name] — use the actual values provided
- Keep total length under 200 words

Output only the email (subject + body), nothing else."""

    message = client.messages.create(
        model=os.getenv("CLAUDE_MODEL", DEFAULT_MODEL),
        max_tokens=600,
        messages=[{"role": "user", "content": prompt}],
    )

    text_blocks = [b.text for b in message.content if getattr(b, "type", "") == "text"]
    if not text_blocks or not text_blocks[0].strip():
        raise ValueError("Claude returned no text for the draft")
    return text_blocks[0].strip()


def get_email_prompt_template() -> str:
    return (
        "Write a professional, concise, and warm {ordinal} follow-up email for a job application.\n\n"
        "Details:\n"
        "- Company: {company}\n"
        "- Role: {role}\n"
        "- Recruiter/Contact: {recruiter_name}\n"
        "- Application Date: {applied_date}\n"
        "- Sender Name: {sender_name}\n"
        "- Additional Context: {notes}\n\n"
        "Requirements:\n"
        "- Subject line on first line starting with \"Subject: \"\n"
        "- 3-4 short paragraphs max\n"
        "- Professional but warm tone\n"
        "- Express genuine interest without being pushy\n"
        "- Mention the specific role and company name\n"
        "- End with a clear call to action\n"
        "- Include a polite sign-off with the sender's name\n"
        "- Do NOT use placeholder brackets like [Your Name] — use the actual values provided\n"
        "- Keep total length under 200 words\n\n"
        "Output only the email (subject + body), nothing else."
    )


def generate_bulk_drafts(applications: list[dict], sender_name: str = "Job Applicant") -> dict[str, str]:
    """Generate AI email drafts for all pending applications. Returns {app_id: draft}."""
    drafts = {}
    for app in applications:
        status = app.get("Status", "")
        email_sent = app.get("Email Sent", "No")
        if status in ("Rejected", "Withdrawn", "Offer Received", "Interview Scheduled"):
            continue
        if email_sent == "Yes":
            continue

        try:
            draft = generate_followup_email(
                company=app.get("Company", ""),
                role=app.get("Role", ""),
                recruiter_name=app.get("Recruiter Name", ""),
                applied_date=app.get("Applied Date", ""),
                follow_up_count=int(app.get("Follow Up Count", 0) or 0),
                notes=app.get("Notes", ""),
                sender_name=sender_name,
            )
            drafts[app.get("ID", "")] = draft
        except Exception as e:
            print(f"Error generating draft for {app.get('Company')}: {e}")

    return drafts
