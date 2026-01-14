# -*- coding: utf-8 -*-
import os
import json
import base64
import requests
import pandas as pd
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from datetime import datetime

# ---------------- GOOGLE SHEET CONFIG -----------------
SPREADSHEET_ID = "1-gAUMbVOio3mTzfDstqjpnQdibP2oYjuF-vhX5UovCw"
RANGE_NAME = "Time_Table_2"

# ---------------- META WHATSAPP CONFIG -----------------
WHATSAPP_TOKEN = os.environ.get("WHATSAPP_TOKEN")
WHATSAPP_PHONE_NUMBER_ID = os.environ.get("WHATSAPP_PHONE_NUMBER_ID")

WHATSAPP_API_URL = (
    f"https://graph.facebook.com/v19.0/{WHATSAPP_PHONE_NUMBER_ID}/messages"
)

# ---------------- SERVICE ACCOUNT -----------------
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
        log_message("❌ No data found.")
        return None

    df = pd.DataFrame(values[1:], columns=values[0])
    log_message(f"✅ Sheet loaded. Rows: {len(df)}")
    return df

# ---------------- SEND WHATSAPP (META API) -----------------
def send_whatsapp(to_phone, message_text):
    headers = {
        "Authorization": f"Bearer {WHATSAPP_TOKEN}",
        "Content-Type": "application/json"
    }

    payload = {
        "messaging_product": "whatsapp",
        "to": to_phone,
        "type": "text",
        "text": {
            "body": message_text
        }
    }

    response = requests.post(
        WHATSAPP_API_URL,
        headers=headers,
        json=payload
    )

    if response.status_code == 200:
        log_message(f"✅ WhatsApp sent to {to_phone}")
    else:
        log_message(
            f"❌ WhatsApp failed to {to_phone}: {response.text}"
        )

# ---------------- PROCESS REMINDERS -----------------
def process_reminders():
    df = read_google_sheet()
    if df is None:
        return

    today_str = datetime.now().strftime("%Y-%m-%d")
    log_message(f"Processing reminders for {today_str}")
    sent_count = 0

    for _, row in df.iterrows():
        if row["Reminder_Date"] != today_str:
            continue

        message = f"""
📢 *Class Reminder*

Dear {row['Customer']},

{row['Message']}

📅 Date: {row['Reminder_Date']}
📘 Course: {row['Course']}
⏰ Time: {row['Session']}

🔗 Zoom Link:
{row['Zoom_link']}

🆔 Meeting ID: {row['Meeting_id']}
🔑 Passcode: {row['Passcode']}

Warm regards,
*New Dimension Academy*
📞 +1 437 967 5082
🌐 www.ndacademy.ca
"""

        send_whatsapp(
            to_phone=row["Phone"],  # Must be E.164 format
            message_text=message.strip()
        )

        sent_count += 1

    log_message(f"🎉 Done. WhatsApp messages sent: {sent_count}")

# ---------------- MAIN -----------------
if __name__ == "__main__":
    process_reminders()
