import json
import os
import sys
import pandas as pd
import requests
from datetime import datetime, timedelta, timezone
from dotenv import load_dotenv
from llm_classifier import configure_openai, classify_response
from email_parser import parse_emails
from excel_writer import  archive_old_no_response_entries, save_to_excel
from report_generator import generate_summary_report
from auth import authenticate_graph, load_config

load_dotenv()

def fetch_job_emails(access_token, user_email, is_ci):
    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    #since = (datetime.now(timezone.utc) - timedelta(days=12)).strftime("%Y-%m-%dT%H:%M:%SZ")
    since = get_last_run()
    print("📅 Fetching emails since:", since)

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
    all_emails = []
    while url:
        response = requests.get(url, headers=headers)
        print("Fetching:", url)
        if response.status_code != 200:
            print("❌ Failed to fetch emails:", response.status_code, response.text)
            break

        data = response.json()
        all_emails.extend(data.get("value", []))
        url = data.get("@odata.nextLink")  # next page if exists

    return all_emails
    

    # response = requests.get(url, headers=headers)
    # print("Fetching from URL:", url)
    # print("Response:", response.status_code, response.text)
    # if response.status_code != 200:
    #     print("❌ Failed to fetch emails:", response.status_code, response.text)
    #     return []

    # return response.json().get("value", [])
LAST_RUN_FILE = "last_run.json"
def get_last_processed_date(excel_file="job_applications.xlsx") -> str:
    """Get the most recent date_applied from Excel, or fallback to 12 days ago."""
    if not os.path.exists(excel_file):
        return (datetime.now(timezone.utc) - timedelta(days=200)).strftime("%Y-%m-%dT00:00:00Z")

    df = pd.read_excel(excel_file)
    if "date_applied" not in df.columns or df.empty:
        return (datetime.now(timezone.utc) - timedelta(days=200)).strftime("%Y-%m-%dT00:00:00Z")

    latest = pd.to_datetime(df["date_applied"], errors="coerce").max()
    if pd.isnull(latest):
        return (datetime.now(timezone.utc) - timedelta(days=200)).strftime("%Y-%m-%dT00:00:00Z")

    return latest.strftime("%Y-%m-%dT00:00:00Z")

def get_last_run() -> str:
    """Return last run time from last_run.json, or fallback to Excel or 12 days ago."""
    if os.path.exists(LAST_RUN_FILE):
        try:
            with open(LAST_RUN_FILE, "r") as f:
                last_run = json.load(f).get("last_run")
                if last_run:
                    return last_run
        except Exception as e:
            print("⚠️ Failed to read last_run.json:", e)

    # Fallback to Excel
    try:
        return get_last_processed_date()
    except Exception as e:
        print("⚠️ Failed to read from Excel:", e)

    # Final fallback
    return (datetime.now(timezone.utc) - timedelta(days=12)).strftime("%Y-%m-%dT%H:%M:%SZ")


def save_last_run(date_str):
    with open(LAST_RUN_FILE, "w") as f:
        json.dump({"last_run": date_str}, f)

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
    #print("🔑 Partial access token:", access_token, "")  # Do NOT log full token

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
        archive_old_no_response_entries()
        generate_summary_report()
        save_last_run(datetime.utcnow().strftime("%Y-%m-%dT00:00:00Z"))

    else:
        print("No job-related emails found.")
