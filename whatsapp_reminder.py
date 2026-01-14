# -*- coding: utf-8 -*-
import os
import json
import base64
import pandas as pd
from datetime import datetime
from twilio.rest import Client
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

# ---------------- CONFIG -----------------
SPREADSHEET_ID = "1-gAUMbVOio3mTzfDstqjpnQdibP2oYjuF-vhX5UovCw"
RANGE_NAME = "Time_Table_2"

# Twilio
TWILIO_ACCOUNT_SID = os.environ["TWILIO_ACCOUNT_SID"]
TWILIO_AUTH_TOKEN = os.environ["TWILIO_AUTH_TOKEN"]
TWILIO_WHATSAPP_FROM = os.environ["TWILIO_WHATSAPP_FROM"]

# Service Account
SERVICE_ACCOUNT_INFO = json.loads(
    base64.b64decode(os.environ["SERVICE_ACCOUNT_JSON"]).decode("utf-8")
)

client = Client(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN)

# ---------------- LOG -----------------
def log(msg):
    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}")

# ---------------- READ SHEET -----------------
def read_google_sheet():
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
        return None

    return pd.DataFrame(values[1:], columns=values[0])

# ---------------- SEND WHATSAPP -----------------
def send_whatsapp(to_number, message):
    client.messages.create(
        from_=TWILIO_WHATSAPP_FROM,
        to=f"whatsapp:{to_number}",
        body=message
    )

# ---------------- PROCESS REMINDERS -----------------
def process_whatsapp_reminders():
    df = read_google_sheet()
    if df is None:
        log("No data found.")
        return

    today = datetime.now().strftime("%Y-%m-%d")
    sent = 0

    for _, row in df.iterrows():
        if row["Reminder_Date"] != today:
            continue

        whatsapp_message = f"""
📢 *Class Reminder – New Dimension Academy*

Dear {row['Customer']},

{row['Message']}

📅 *Date:* {row['Reminder_Date']}
📘 *Course:* {row['Course']}
⏰ *Time:* {row['Session']}

🔗 *Zoom Link:*
{row['Zoom_link']}

Warm regards,
*New Dimension Academy*
📞 +1 437 967 5082
🌐 www.ndacademy.ca

_Expanding Minds, Unlocking New Dimensions_
"""

        send_whatsapp(row["Phone"], whatsapp_message)
        sent += 1
        log(f"WhatsApp sent to {row['Phone']}")

    log(f"✅ Done — {sent} WhatsApp reminder(s) sent.")

# ---------------- MAIN -----------------
if __name__ == "__main__":
    process_whatsapp_reminders()
