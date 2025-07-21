import json
from msal import PublicClientApplication
import requests
from datetime import datetime, timedelta
from datetime import timezone
from llm_classifier import configure_openai, classify_response
from email_parser import parse_emails
from excel_writer import save_to_excel
from report_generator import generate_summary_report

CONFIG_FILE = "config.json"
TOKEN_FILE = "tokens.json"

def load_config():
    with open(CONFIG_FILE, "r") as f:
        return json.load(f)

def save_token(token):
    with open(TOKEN_FILE, "w") as f:
        json.dump(token, f)

def load_token():
    try:
        with open(TOKEN_FILE, "r") as f:
            return json.load(f)
    except FileNotFoundError:
        return None

def authenticate_graph(config):
    app = PublicClientApplication(
        config["client_id"],
        authority=f"https://login.microsoftonline.com/{config['tenant_id']}"
    )

    accounts = app.get_accounts()
    result = None

    if accounts:
        result = app.acquire_token_silent(config["scopes"], account=accounts[0])

    if not result:
        try:
            flow = app.initiate_device_flow(scopes=config["scopes"])
        except Exception as e:
            raise Exception(f"Device flow initiation failed: {str(e)}")

        if "user_code" not in flow:
            raise Exception(f"Device flow failed. Response: {flow}")

        print("Go to:", flow["verification_uri"])
        print("Enter the code:", flow["user_code"])

        result = app.acquire_token_by_device_flow(flow)

    if "access_token" in result:
        save_token(result)
        return result["access_token"]
    else:
        raise Exception(f"Authentication failed: {result.get('error_description', result)}")


def fetch_job_emails(access_token):
    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    
    since = (datetime.now(timezone.utc) - timedelta(days=10)).strftime("%Y-%m-%dT%H:%M:%SZ")


    url = (
        "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages"
        f"?$filter=receivedDateTime ge {since}&$top=50"
        "&$select=subject,bodyPreview,receivedDateTime,from"
    )

    response = requests.get(url, headers=headers)

    if response.status_code != 200:
        print("Failed to fetch emails:", response.status_code, response.text)
        return []

    return response.json().get("value", [])

if __name__ == "__main__":
    config = load_config()
    
    configure_openai(config["openai_api_key"])  # ← THIS MUST BE HERE

    access_token = authenticate_graph(config)
    emails = fetch_job_emails(access_token)

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
