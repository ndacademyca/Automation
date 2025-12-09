# -*- coding: utf-8 -*-
import os
import json
import base64
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from datetime import datetime

# ---------------- GOOGLE SHEET CONFIG -----------------
SPREADSHEET_ID = "14VaN5r4mW94J8nPS6P7C40jHh7oULnyv0438rcB2k0E"
RANGE_NAME = "Sheet1"

# ---------------- EMAIL CONFIG (Gmail) -----------------
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465
EMAIL_USER = os.environ.get("EMAIL_USER")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD")  # App password

# ---------------- SERVICE ACCOUNT FROM SECRET -----------------
if "SERVICE_ACCOUNT_JSON" not in os.environ:
    raise ValueError("SERVICE_ACCOUNT_JSON environment variable is not set!")

SERVICE_ACCOUNT_INFO = json.loads(
    base64.b64decode(os.environ["SERVICE_ACCOUNT_JSON"]).decode("utf-8")
)

# ---------------- LOG FUNCTION -----------------
def log_message(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}")

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
        df = pd.DataFrame(values[1:], columns=values[0])
        log_message(f"✅ Google Sheet read successfully. Rows: {len(df)}")
        return df
    except Exception as e:
        log_message(f"❌ Failed to read Google Sheet: {e}")
        return None

# ---------------- SEND EMAIL -----------------
def send_email(recipient, subject, body):
    try:
        msg = MIMEMultipart()
        msg["From"] = EMAIL_USER
        msg["To"] = recipient
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "html"))

        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=context) as server:
            server.login(EMAIL_USER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_USER, recipient, msg.as_string())

        log_message(f"✅ Email sent to {recipient}")
    except Exception as e:
        log_message(f"❌ Failed to send email to {recipient}: {e}")

# ---------------- PROCESS REMINDERS -----------------
def process_reminders():
    df = read_google_sheet()
    if df is None:
        return "No data to process."

    today_str = datetime.now().strftime("%Y-%m-%d")
    log_message(f"Processing reminders for {today_str}")
    sent_count = 0

    for _, row in df.iterrows():
        if row["Reminder_Date"] != today_str:
            continue

        body = f"""
        <p>Dear {row['Customer']},</p>
        <p>{row['Message']}</p>
        <p><b>Class Date:</b> {row['Reminder_Date']}<br>
        <b>Course:</b> {row['Course']}<br>
        <b>Class Time:</b> {row['Session']}</p>
        <p><img src="https://raw.githubusercontent.com/ndacademyca/images/main/Whatsapp_notification.png"
                width="650"></p>
        <p><b>Zoom link:</b> {row['Zoom_link']}</p>
        <p>Warm regards,<br>New Dimension Academy</p>
        """

        send_email(row["Email"], f"Reminder for {row['Customer']}", body)
        sent_count += 1

    log_message(f"🎉 All reminders processed. Total emails sent: {sent_count}")
    return f"Done — {sent_count} reminder(s) sent."

# ---------------- MAIN ENTRY POINT -----------------
if __name__ == "__main__":
    process_reminders()
