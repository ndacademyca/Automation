# -*- coding: utf-8 -*-

import os
import json
import base64
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timezone
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

# ---------------- CONFIGURATION -----------------
SPREADSHEET_ID = "1kfrSDM1c8Z9MBtI85IjDjJABL2KXuN1k8yPyAfWKh0U"
RANGE_NAME = "Current_Report"

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")  # App password

HEADER_IMAGE_URL = os.getenv("HEADER_IMAGE_URL", "")
FOOTER_IMAGE_URL = os.getenv("FOOTER_IMAGE_URL", "")

# ---------------- SERVICE ACCOUNT -----------------
if "SERVICE_ACCOUNT_JSON" not in os.environ:
    raise ValueError("SERVICE_ACCOUNT_JSON environment variable is not set!")

# Decode base64 JSON (works with GitHub Actions secret)
SERVICE_ACCOUNT_INFO = json.loads(
    base64.b64decode(os.environ["SERVICE_ACCOUNT_JSON"]).decode("utf-8")
)

# ---------------- LOG FUNCTION -----------------
def log_message(message: str):
    ts = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
    print(f"[{ts}] {message}")

# ---------------- READ GOOGLE SHEET -----------------
def read_google_sheet():
    log_message("📌 read_google_sheet() called")
    try:
        creds = Credentials.from_service_account_info(
            SERVICE_ACCOUNT_INFO,
            scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"]
        )
        service = build("sheets", "v4", credentials=creds)
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=RANGE_NAME
        ).execute()
        values = result.get("values", [])
        if not values:
            log_message("❌ No data found in Google Sheet.")
            return None

        df = pd.DataFrame(values[1:], columns=[c.strip() for c in values[0]])
        df = df.loc[:, df.columns != ""]  # remove empty columns
        df = df.fillna("")  # replace NaN with empty string
        log_message(f"✅ Google Sheet read successfully. Rows: {len(df)}")
        log_message(f"[📄] Columns detected: {list(df.columns)}")
        return df

    except Exception as e:
        log_message(f"❌ Failed to read Google Sheet: {e}")
        return None

# ---------------- BUILD EMAIL -----------------
def build_email(row):
    return f"""
    <!DOCTYPE html>
    <html>
    <head><meta charset="UTF-8"><title>Student Progress Report</title></head>
    <body style="margin:0;padding:0;background:#f4f6f8;font-family:Arial,sans-serif">

    <table width="100%" cellpadding="0" cellspacing="0">
    <tr><td align="center">

        <table width="600" style="background:#fff;border-radius:8px;overflow:hidden;border-collapse:collapse">

            <!-- Header Image -->
            <tr><td><img src="{HEADER_IMAGE_URL}" width="600" style="display:block;width:100%"></td></tr>

            <!-- Student Info -->
            <tr>
                <td style="padding:20px;text-align:center;background:#f9fafb">
                    <h2 style="margin:0">Student Progress Report</h2>
                    <p style="margin:6px 0 0; font-size:14px; color:#7f8c8d;">{row.get('Course_Month','')} {row.get('Course_Year','')}</p>
                </td>
            </tr>

            <tr><td style="padding:20px">
                <table width="100%" cellpadding="8" cellspacing="0" style="border-collapse:collapse">
                    <tr><td style="background:#eef2f5"><strong>Student Name</strong></td><td>{row.get('Student_Name','')}</td></tr>
                    <tr><td style="background:#eef2f5"><strong>Course</strong></td><td>{row.get('Course','')}</td></tr>
                    <tr><td style="background:#eef2f5"><strong>Level</strong></td><td>{row.get('Level','')}</td></tr>
                    <tr><td style="background:#eef2f5"><strong>Teacher</strong></td><td>{row.get('Teacher','')}</td></tr>
                </table>
            </td></tr>

            <!-- Cognitive Goals -->
            <tr><td style="padding:20px">
                <table width="100%" cellpadding="10" cellspacing="0" style="border-collapse:collapse">
                    <tr><td style="background:#34495e; color:#fff"><strong>Cognitive Goals</strong></td></tr>
                    <tr><td style="border:1px solid #e0e0e0;line-height:1.6">{row.get('Cognitive_Goals','')}</td></tr>
                </table>
            </td></tr>

            <!-- Teacher Comments -->
            <tr><td style="padding:20px">
                <table width="100%" cellpadding="10" cellspacing="0" style="border-collapse:collapse">
                    <tr><td style="background:#34495e; color:#fff"><strong>Teacher's Comments</strong></td></tr>
                    <tr><td style="border:1px solid #e0e0e0;line-height:1.6">{row.get("Teacher's_Comments",'')}</td></tr>
                </table>
            </td></tr>

            <!-- General Comment -->
            <tr><td style="padding:20px">
                <table width="100%" cellpadding="10" cellspacing="0" style="border-collapse:collapse">
                    <tr><td style="background:#34495e; color:#fff"><strong>General Comment</strong></td></tr>
                    <tr><td style="border:1px solid #e0e0e0;line-height:1.6">{row.get('General_Comment','')}</td></tr>
                </table>
            </td></tr>

            <!-- Footer -->
            <tr><td style="padding:20px;text-align:center;font-size:13px;color:#7f8c8d">
                Report Date: {row.get('Report_Date','')}<br>
                Contact: {row.get('Teacher','')} – {row.get('Teacher_Email','')}
            </td></tr>

            <!-- Footer Image -->
            <tr><td><img src="{FOOTER_IMAGE_URL}" width="600" style="display:block;width:100%"></td></tr>

        </table>

    </td></tr>
    </table>

    </body>
    </html>
    """

# ---------------- SEND EMAIL -----------------
def send_email(to_email, teacher_email, subject, body):
    try:
        msg = MIMEMultipart()
        msg["From"] = f"New Dimension Academy <{EMAIL_USER}>"
        msg["To"] = to_email
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "html"))

        bcc_list = ["alhuraibia@gmail.com", "dalmaznaee@gmail.com"]
        if teacher_email.strip():
            bcc_list.append(teacher_email)

        recipients = [to_email] + bcc_list
        msg["Bcc"] = ", ".join(bcc_list)

        context = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
        context.login(EMAIL_USER, EMAIL_PASSWORD)
        context.sendmail(EMAIL_USER, recipients, msg.as_string())
        context.quit()

        log_message(f"✅ Email sent → TO: {to_email} | BCC: {', '.join(bcc_list)}")
    except Exception as e:
        log_message(f"❌ Failed to send email to {to_email}: {e}")

# ---------------- PROCESS REPORTS -----------------
def process_reminders():
    df = read_google_sheet()
    if df is None:
        log_message("No data to process")
        return

    today_str = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    log_message(f"📅 Processing reports for {today_str}")

    sent_count = 0
    for _, row in df.iterrows():
        report_date = str(row.get("Report_Date",'')).split(" ")[0]
        if report_date != today_str:
            continue

        log_message(f"📨 Sending report to {row.get('Student_Email','')}")
        email_body = build_email(row)
        subject = f"{row.get('Course_Month','')} {row.get('Course_Year','')} {row.get('Course','')} Progress Report for {row.get('Student_Name','')}"

        send_email(
            to_email=row.get("Student_Email",''),
            teacher_email=row.get("Teacher_Email",''),
            subject=subject,
            body=email_body
        )
        sent_count += 1

    log_message(f"🎉 Completed. Emails sent: {sent_count}")

# ---------------- MAIN -----------------
if __name__ == "__main__":
    process_reminders()
