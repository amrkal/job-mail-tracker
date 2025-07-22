import json
import os
from msal import PublicClientApplication
from some_module import load_config  # Assumes you defined load_config there using os.environ

TOKEN_FILE = "tokens.json"

def load_config():
    return {
        "client_id": os.environ.get("CLIENT_ID"),
        "tenant_id": os.environ.get("TENANT_ID"),
        #"client_secret": os.environ.get("CLIENT_SECRET"),
        "scopes": os.environ.get("SCOPES", "Mail.Read").split(),
        "output_excel": os.environ.get("OUTPUT_EXCEL", "job_applications.xlsx"),
        "report_output_folder": os.environ.get("REPORT_OUTPUT_FOLDER", "reports")
    }

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
        flow = app.initiate_device_flow(scopes=config["scopes"])
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
