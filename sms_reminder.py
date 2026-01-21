# -*- coding: utf-8 -*-
import os
import json
import base64
from datetime import datetime
import pandas as pd
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from twilio.rest import Client

# ---------------- GOOGLE SHEET CONFIG -----------------
SPREADSHEET_ID = "1-gAUMbVOio3mTzfDstqjpnQdibP2oYjuF-vhX5UovCw"
RANGE_NAME = "Time_Table_2"

# ---------------- TWILIO SMS CONFIG -----------------
TWILIO_ACCOUNT_SID = os.environ.get("TWILIO_ACCOUNT_SID")
TWILIO_AUTH_TOKEN = os.environ.get("TWILIO_AUTH_TOKEN")
TWILIO_FROM_NUMBER = os.environ.get("TWILIO_FROM_NUMBER")

if not all([TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN, TWILIO_FROM_NUMBER]):
    raise ValueError("❌ Twilio environment variables are missing")

twilio_client = Client(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN)

# ---------------- SERVICE ACCOUNT -----------------
if "SERVICE_ACCOUNT_JSON" not in os.environ:
    raise ValueError("SERVICE_ACCOUNT_JSON is missing")

SERVICE_ACCOUNT_INFO = json.loads(
    base64.b64decode(os.environ["SERVICE_ACCOUNT_JSON"]).decode("utf-8")
)

# ---------------- LOG FUNCTION -----------------
def log_message(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}")

# ---------------- READ GOOGLE SHEET -----------------
# def read_google_sheet():
#     log_message("📌 read_google_sheet() called")

#     creds = Credentials.from_service_account_info(
#         SERVICE_ACCOUNT_INFO,
#         scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"]
#     )

#     service = build("sheets", "v4", credentials=creds)
#     result = service.spreadsheets().values().get(
#         spreadsheetId=SPREADSHEET_ID,
#         range=RANGE_NAME
#     ).execute()

#     values = result.get("values", [])
#     if not values:
#         log_message("❌ No data found in Google Sheet.")
#         return None

#     df = pd.DataFrame(values[1:], columns=values[0])
#     log_message(f"✅ Sheet loaded. Rows: {len(df)}")
#     return df

def read_google_sheet(retries=3, timeout=60):
    log_message("📌 read_google_sheet() called")

    creds = Credentials.from_service_account_info(
        SERVICE_ACCOUNT_INFO,
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"]
    )

    service = build(
        "sheets",
        "v4",
        credentials=creds,
        cache_discovery=False
    )

    for attempt in range(1, retries + 1):
        try:
            result = service.spreadsheets().values().get(
                spreadsheetId=SPREADSHEET_ID,
                range=RANGE_NAME
            ).execute(num_retries=3)

            values = result.get("values", [])
            if not values:
                log_message("❌ No data found in Google Sheet.")
                return None

            df = pd.DataFrame(values[1:], columns=values[0])
            log_message(f"✅ Sheet loaded. Rows: {len(df)}")
            return df

        except HttpError as e:
            log_message(f"⚠️ Google API error (attempt {attempt}): {e}")
        except TimeoutError:
            log_message(f"⏱ Timeout (attempt {attempt})")
        except Exception as e:
            log_message(f"❌ Unexpected error: {e}")

        if attempt < retries:
            time.sleep(5)

    log_message("❌ Failed to read Google Sheet after retries")
    return None


# ---------------- SEND SMS -----------------
def send_sms(
    to_phone,
    customer,
    course,
    #class_date,
    class_time#,
    #zoom_link
):
    message_body = (
        f"Hello {customer},\n"
        f"You have a class Today.\n"
        f"Course: {course}\n"
        #f"Date: {class_date}\n"
        f"Time: {class_time}\n"
        #f"Zoom: {zoom_link}\n\n"
        #f"See you soon!"
        f"Let’s learn and have fun😊\n"
        f"New Dimension Academy"
    )

    try:
        message = twilio_client.messages.create(
            body=message_body,
            from_=TWILIO_FROM_NUMBER,
            to=to_phone
        )

        log_message(f"✅ SMS sent to {to_phone} (SID: {message.sid})")

    except Exception as e:
        log_message(f"❌ Failed to send SMS to {to_phone}: {str(e)}")

# ---------------- PROCESS REMINDERS -----------------
def process_reminders():
    df = read_google_sheet()
    if df is None:
        return

    today_str = datetime.now().strftime("%Y-%m-%d")
    log_message(f"📅 Processing reminders for {today_str}")

    sent_count = 0

    for _, row in df.iterrows():
        if row["Reminder_Date"] != today_str:
            continue

        send_sms(
            to_phone=row["Phone"],
            customer=row["Customer"],
            course=row["Course"],
            #class_date=row["Reminder_Date"],
            class_time=row["Session"],
            #zoom_link=row["Zoom_link"]
        )

        sent_count += 1

    log_message(f"🎉 Done. SMS messages sent: {sent_count}")

# ---------------- MAIN -----------------
if __name__ == "__main__":
    process_reminders()
