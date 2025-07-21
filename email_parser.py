import re
from datetime import datetime

def normalize_company(email_address: str, subject: str) -> str:
    domain = email_address.split("@")[-1].split(".")[0]
    if "comeet" in email_address:
        match = re.search(r'at ([\w\-\.]+)', subject, re.IGNORECASE)
        return match.group(1).strip() if match else domain
    return domain

def extract_job_title(subject: str) -> str:
    match = re.search(r"for (.+?) position", subject, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    match = re.search(r"מועמדות למשרת\s*(.+?)\b", subject, re.IGNORECASE)
    return match.group(1).strip() if match else "Unknown"

def parse_emails(messages: list) -> list:
    parsed = []
    seen_threads = set()

    for msg in messages:
        thread_id = msg["subject"].lower().strip()
        if thread_id in seen_threads:
            continue
        seen_threads.add(thread_id)

        sender = msg.get("from", {}).get("emailAddress", {}).get("address", "")
        subject = msg.get("subject", "")
        preview = msg.get("bodyPreview", "")
        date = msg.get("receivedDateTime", "")
        parsed.append({
            "company": normalize_company(sender, subject),
            "job_title": extract_job_title(subject),
            "date_applied": datetime.strptime(date, "%Y-%m-%dT%H:%M:%SZ").strftime("%Y-%m-%d"),
            "subject": subject,
            "email": sender,
            "thread_id": thread_id,
            "preview": preview,
        })

    return parsed
