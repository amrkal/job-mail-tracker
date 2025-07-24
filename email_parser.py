import re
from datetime import datetime
from collections import defaultdict


def normalize_company(email_address: str, subject: str, preview: str = "") -> str:
    if not email_address or "@" not in email_address:
        return "Unknown"

    full_domain = email_address.split("@")[-1]
    domain_parts = full_domain.split(".")
    domain = domain_parts[0].lower()

    known_platforms = ["comeet", "greenhouse", "linkedin", "workflow", "mail", "smartrecruiters", "myworkday", "canditech", "sparkhire"]

    # If it's a known hiring platform, try to extract from subject or preview
    if any(p in full_domain for p in known_platforms):
        # Match from "at Company", "from Company", etc.
        combined = f"{subject} {preview}"
        match = re.search(r"\b(?:at|from|on behalf of)\s+([A-Z][\w&\-\. ]+)", combined, re.IGNORECASE)
        if match:
            return match.group(1).strip(" .,-").title()

    # Otherwise fallback to domain name
    if domain in known_platforms:
        if len(domain_parts) > 1:
            company_guess = domain_parts[-2]
            return company_guess.title()
        return "Unknown"

    return domain.title()


def extract_job_title(subject: str, preview: str = "") -> str:
    combined = f"{subject} {preview}"

    patterns = [
        r"applying for (?:the )?(.+?)(?: position| role| job| at|\.|$)",
        r"application for (?:the )?(.+?)(?: position| role| job| at|\.|$)",
        r"for (?:the )?(.+?) position",
        r"position[:\-] (.+?)(?:\.|$)",
        r"role[:\-] (.+?)(?:\.|$)",
        r"מועמדות למשרת\s*(.+?)\b",
        r"position at\s+([A-Z][\w\- ]+)",  # for cases like "Software Engineer Position at X"
        r"invitation to (.+?) interview",
        r"technical exam and guidelines for (.+?)",
        r"joining our (.+?) program",
    ]

    for pattern in patterns:
        match = re.search(pattern, combined, re.IGNORECASE)
        if match:
            job = match.group(1).strip(" .,-")
            if len(job.split()) <= 8:  # avoid overly long matches
                return job.title()

    return "Unknown"





def parse_emails(messages: list) -> list:
    parsed_by_sender = defaultdict(list)

    for msg in messages:
        sender = msg.get("from", {}).get("emailAddress", {}).get("address", "").lower()
        subject = msg.get("subject", "")
        preview = msg.get("bodyPreview", "")
        date_str = msg.get("receivedDateTime", "")
        if not sender or not subject or not date_str:
            continue

        try:
            received_dt = datetime.strptime(date_str, "%Y-%m-%dT%H:%M:%SZ")
        except ValueError:
            continue

        parsed_by_sender[sender].append((received_dt, {
            "company": normalize_company(sender, subject),
            "job_title": extract_job_title(subject),
            "date_applied": received_dt.strftime("%Y-%m-%d"),
            "subject": subject,
            "email": sender,
            "thread_id": subject.lower().strip(),
            "preview": preview,
        }))

    # Keep only the latest email from each sender
    latest_per_sender = [
        sorted(emails, key=lambda x: x[0], reverse=True)[0][1]  # take newest
        for emails in parsed_by_sender.values()
    ]

    return latest_per_sender