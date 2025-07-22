import json
import os
import sys
import requests
from datetime import datetime, timedelta, timezone
from dotenv import load_dotenv
from llm_classifier import configure_openai, classify_response
from email_parser import parse_emails
from excel_writer import save_to_excel
from report_generator import generate_summary_report
from auth import authenticate_graph, load_config

load_dotenv()

def fetch_job_emails(access_token, user_email, is_ci):
    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    since = (datetime.now(timezone.utc) - timedelta(days=10)).strftime("%Y-%m-%dT%H:%M:%SZ")

    if is_ci:
        # App-only auth (client credentials)
        url = (
            f"https://graph.microsoft.com/v1.0/users/{user_email}/mailFolders/inbox/messages"
            f"?$filter=receivedDateTime ge {since}&$top=50"
            "&$select=subject,bodyPreview,receivedDateTime,from"
        )
    else:
        # Delegated auth (device code flow)
        url = (
            f"https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages"
            f"?$filter=receivedDateTime ge {since}&$top=50"
            "&$select=subject,bodyPreview,receivedDateTime,from"
        )
    

    response = requests.get(url, headers=headers)
    print("Fetching from URL:", url)
    print("Response:", response.status_code, response.text)
    if response.status_code != 200:
        print("❌ Failed to fetch emails:", response.status_code, response.text)
        return []

    return response.json().get("value", [])

if __name__ == "__main__":
    openai_key = os.environ["OPENAI_API_KEY"]
    configure_openai(openai_key)

    config = load_config()
    is_ci = os.environ.get("CI") == "true"

    print("Loaded config keys:", config.keys())
    if is_ci and not config.get("client_secret"):
        print("❌ CLIENT_SECRET is missing in CI environment.")
        sys.exit(1)

    access_token = authenticate_graph(config)
    print("🔑 Partial access token:", access_token, "")  # Do NOT log full token

    emails = fetch_job_emails(access_token, config["user_email"], is_ci)
    print(f"Fetched {len(emails)} emails.")

    parsed = []
    for email in emails:
        metadata = parse_emails([email])[0]
        metadata["response_type"] = classify_response(metadata["subject"], email.get("bodyPreview", ""))
        parsed.append(metadata)
        print(metadata)

    if parsed:
        save_to_excel(parsed)
        generate_summary_report()
    else:
        print("No job-related emails found.")
