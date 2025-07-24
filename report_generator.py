import pandas as pd
from datetime import datetime
import os
import pdfkit

EXCEL_FILE = "job_applications.xlsx"
REPORTS_DIR = "reports"


path_wkhtmltopdf = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)



def generate_summary_report():
    if not os.path.exists(EXCEL_FILE):
        print("❌ Excel file not found.")
        return

    df = pd.read_excel(EXCEL_FILE)

    # Ensure required column exists
    if "response_type" not in df.columns:
        print("❌ Missing 'response_type' column.")
        return

    # Filter: only show Applied / No Reply Yet
    df_pending = df[df["response_type"].isin(["Applied", "No Reply Yet"])]
    df_pending = df_pending.sort_values(by=["response_type", "date_applied"])

    # Format date
    today_str = datetime.now().strftime("%Y-%m-%d")
    os.makedirs(REPORTS_DIR, exist_ok=True)

    # CSV Export
    csv_path = os.path.join(REPORTS_DIR, f"job_followup_{today_str}.csv")
    df_pending.to_csv(csv_path, index=False)

    # Optional: PDF Export via HTML
    try:
        import pdfkit
        html = df_pending.to_html(index=False)
        pdf_path = os.path.join(REPORTS_DIR, f"job_followup_{today_str}.pdf")
        #pdfkit.from_string(html, pdf_path, configuration=config)
        pdfkit.from_string(html, pdf_path, configuration=config)
        print(f"📄 Report saved as: {csv_path} and {pdf_path}")
    except ImportError:
        print(f"✅ CSV report saved: {csv_path} (PDF skipped - install `pdfkit` for PDF output)")
