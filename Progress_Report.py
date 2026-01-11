import os
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timezone
import gspread
from google.oauth2.service_account import Credentials


# =============================
# CONFIGURATION
# =============================
SPREADSHEET_NAME = "Progress_Report"
WORKSHEET_NAME = "Sheet1"

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")


# =============================
# GOOGLE SHEETS
# =============================
def read_google_sheet():
    print("[📌] read_google_sheet() called")

    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_info(
        eval(os.getenv("SERVICE_ACCOUNT_JSON")),
        scopes=scopes
    )

    client = gspread.authorize(creds)
    sheet = client.open(SPREADSHEET_NAME).worksheet(WORKSHEET_NAME)

    data = sheet.get_all_records()
    df = pd.DataFrame(data)

    # 🔹 Drop empty columns
    df = df.loc[:, df.columns != ""]

    print(f"[✅] Google Sheet read successfully. Rows: {len(df)}")
    print(f"[📄] Columns detected: {list(df.columns)}")

    return df


# =============================
# EMAIL TEMPLATE
# =============================
def build_email(row):
    return f"""
    <html>
    <body style="margin:0;padding:0;background:#f4f4f4;font-family:Arial,sans-serif">
        <table width="100%" cellpadding="0" cellspacing="0">
            <tr>
                <td align="center">
                    <table width="600" style="background:#ffffff;border-collapse:collapse">

                        <!-- Header Image -->
                        <tr>
                            <td>
                                <img src="https://example.com/header.png" width="600" style="display:block">
                            </td>
                        </tr>

                        <!-- Student Info -->
                        <tr>
                            <td style="padding:20px">
                                <h2>Progress Report</h2>
                                <p><strong>Student:</strong> {row.get('Student_Name','')}</p>
                                <p><strong>Course:</strong> {row.get('Course','')}</p>
                                <p><strong>Level:</strong> {row.get('Level','')}</p>
                                <p><strong>Teacher:</strong> {row.get('Teacher','')}</p>
                            </td>
                        </tr>

                        <!-- Cognitive Goals -->
                        <tr>
                            <td style="padding:20px">
                                <strong>Cognitive Goals</strong><br>
                                {row.get('Cognitive_Goals','')}
                            </td>
                        </tr>

                        <!-- Teacher Comments -->
                        <tr>
                            <td style="padding:20px">
                                <strong>Teacher Comments</strong><br>
                                {row.get("Teacher's_Comments",'')}
                            </td>
                        </tr>

                        <!-- General Comment -->
                        <tr>
                            <td style="padding:20px">
                                <strong>General Comment</strong><br>
                                {row.get('General_Comment','')}
                            </td>
                        </tr>

                        <!-- Footer Image -->
                        <tr>
                            <td>
                                <img src="https://example.com/footer.png" width="600" style="display:block">
                            </td>
                        </tr>

                    </table>
                </td>
            </tr>
        </table>
    </body>
    </html>
    """


# =============================
# SEND EMAIL
# =============================
def send_email(to_email, subject, html_body):
    msg = MIMEMultipart("alternative")
    msg["From"] = EMAIL_USER
    msg["To"] = to_email
    msg["Subject"] = subject

    msg.attach(MIMEText(html_body, "html"))

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(EMAIL_USER, EMAIL_PASSWORD)
        server.send_message(msg)


# =============================
# MAIN PROCESS
# =============================
def process_reminders():
    df = read_google_sheet()

    today_str = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    print(f"[📅] Processing reminders for {today_str}")

    for _, row in df.iterrows():
        report_date = str(row.get("Report_Date", ""))[:10]

        if report_date != today_str:
            continue

        print(f"[📨] Sending report to {row.get('Student_Email')}")

        email_body = build_email(row)

        send_email(
            to_email=row.get("Student_Email"),
            subject="Student Progress Report",
            html_body=email_body
        )

        print("[✅] Email sent successfully")


# =============================
# ENTRY POINT
# =============================
if __name__ == "__main__":
    process_reminders()
