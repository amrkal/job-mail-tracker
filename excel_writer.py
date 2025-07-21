import pandas as pd
import os

EXCEL_FILE = "job_applications.xlsx"

COLUMNS = [
    "company",
    "job_title",
    "date_applied",
    "response_type",
    "subject",
    "email",
    "thread_id"
]

def save_to_excel(records):
    if os.path.exists(EXCEL_FILE):
        df_existing = pd.read_excel(EXCEL_FILE)
    else:
        df_existing = pd.DataFrame(columns=COLUMNS)

    # Deduplicate by thread_id
    if "thread_id" not in df_existing.columns:
        print("⚠️ Excel file missing 'thread_id' column. Recreating file.")
        df_existing = pd.DataFrame(columns=COLUMNS)

    new_records = []
    existing_ids = set(df_existing["thread_id"].astype(str).str.lower())

    for r in records:
        if r["thread_id"].lower() not in existing_ids:
            new_records.append(r)

    if not new_records:
        print("No new records to write.")
        return

    df_new = pd.DataFrame(new_records)
    df_all = pd.concat([df_existing, df_new], ignore_index=True)
    df_all.to_excel(EXCEL_FILE, index=False)
    print(f"✔️ Saved {len(new_records)} new records to {EXCEL_FILE}")
