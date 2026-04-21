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
SPREADSHEET_ID = "1mhTdW15u6E-jODDpXdlJjZohVU2NHbmzF2R8TZEpIls"
RANGE_NAME = "Invoices"

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")  # App password

HEADER_IMAGE_URL = os.getenv("HEADER_IMAGE_URL", "")
FOOTER_IMAGE_URL = os.getenv("FOOTER_IMAGE_URL", "")

# ---------------- SERVICE ACCOUNT -----------------
if "SERVICE_ACCOUNT_JSON" not in os.environ:
    raise ValueError("SERVICE_ACCOUNT_JSON environment variable is not set!")

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
        df = df.loc[:, df.columns != ""]
        df = df.fillna("")
        log_message(f"✅ Google Sheet read successfully. Rows: {len(df)}")
        log_message(f"[📄] Columns detected: {list(df.columns)}")
        return df

    except Exception as e:
        log_message(f"❌ Failed to read Google Sheet: {e}")
        return None

# ---------------- FORMAT CURRENCY -----------------
def fmt_currency(value):
    """Format a value as currency, return as-is if not numeric."""
    try:
        return f"{float(value):,.2f}"
    except (ValueError, TypeError):
        return str(value) if value else "—"

# ---------------- BUILD DISCOUNT ROWS -----------------
def build_discount_rows(row):
    """Build HTML rows for any non-empty discount types."""
    html = ""
    for i in range(1, 4):
        discount = str(row.get(f"Discount Type {i}", "")).strip()
        if discount:
            html += f"""
            <tr>
                <td colspan="3" style="padding:10px 12px;border:1px solid #e0e0e0;color:#c0392b">
                    Discount — {discount}
                </td>
                <td style="padding:10px 12px;border:1px solid #e0e0e0;color:#c0392b;text-align:right">
                    − {fmt_currency(row.get('Total Discount', ''))}
                </td>
            </tr>"""
            break  # Total Discount is a single combined value; show once with all discount labels
    return html

# ---------------- BUILD EMAIL -----------------
def build_email(row):
    discount_rows = build_discount_rows(row)

    # Determine if any discount exists
    has_discount = any(
        str(row.get(f"Discount Type {i}", "")).strip()
        for i in range(1, 4)
    )

    subtotal_row = f"""
        <tr style="background:#f9fafb">
            <td colspan="3" style="padding:10px 12px;border:1px solid #e0e0e0;font-weight:bold">Subtotal</td>
            <td style="padding:10px 12px;border:1px solid #e0e0e0;text-align:right;font-weight:bold">
                {fmt_currency(row.get('Amount', ''))}
            </td>
        </tr>""" if has_discount else ""

    return f"""
    <!DOCTYPE html>
    <html>
    <head><meta charset="UTF-8"><title>Invoice</title></head>
    <body style="margin:0;padding:0;background:#f4f6f8;font-family:Arial,sans-serif">

    <table width="100%" cellpadding="0" cellspacing="0">
    <tr><td align="center" style="padding:30px 0">

        <table width="620" style="background:#fff;border-radius:10px;overflow:hidden;border-collapse:collapse;box-shadow:0 2px 8px rgba(0,0,0,0.08)">

            <!-- Header Image -->
            {"<tr><td><img src='" + HEADER_IMAGE_URL + "' width='620' style='display:block;width:100%'></td></tr>" if HEADER_IMAGE_URL else ""}

            <!-- Invoice Title Bar -->
            <tr>
                <td style="padding:24px 28px;background:#1a2e44;color:#fff">
                    <table width="100%" cellpadding="0" cellspacing="0">
                        <tr>
                            <td>
                                <h2 style="margin:0;font-size:22px;letter-spacing:1px">INVOICE</h2>
                                <p style="margin:4px 0 0;font-size:13px;color:#aab8c4">
                                    {row.get('Service Month', '')} {row.get('Service Year', '')}
                                </p>
                            </td>
                            <td align="right">
                                <p style="margin:0;font-size:13px;color:#aab8c4">Invoice No.</p>
                                <p style="margin:4px 0 0;font-size:18px;font-weight:bold;color:#f0c040">
                                    #{row.get('Invoice Numeber', row.get('Invoice Number', ''))}
                                </p>
                                <p style="margin:6px 0 0;font-size:12px;color:#aab8c4">
                                    Date: {row.get('Invoice Date', '')}
                                </p>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>

            <!-- Billed To -->
            <tr>
                <td style="padding:20px 28px;background:#f0f4f8;border-bottom:2px solid #dce3ea">
                    <p style="margin:0 0 6px;font-size:11px;text-transform:uppercase;color:#7f8c8d;letter-spacing:1px">Billed To</p>
                    <p style="margin:0;font-size:16px;font-weight:bold;color:#1a2e44">{row.get('Student', '')}</p>
                    <p style="margin:4px 0 0;font-size:13px;color:#555">{row.get('Customer Email', '')}</p>
                    <p style="margin:2px 0 0;font-size:13px;color:#555">{row.get('Customer Mobile No.', row.get('Customer Mobile No', ''))}</p>
                </td>
            </tr>

            <!-- Course Details -->
            <tr>
                <td style="padding:20px 28px">
                    <p style="margin:0 0 12px;font-size:11px;text-transform:uppercase;color:#7f8c8d;letter-spacing:1px">Course Details</p>
                    <table width="100%" cellpadding="8" cellspacing="0" style="border-collapse:collapse;font-size:14px">
                        <tr>
                            <td width="50%" style="padding:8px 10px;background:#f9fafb;border:1px solid #e8ecef"><strong>Course</strong></td>
                            <td style="padding:8px 10px;border:1px solid #e8ecef">{row.get('Course_', row.get('Course', ''))}</td>
                        </tr>
                        <tr>
                            <td style="padding:8px 10px;background:#f9fafb;border:1px solid #e8ecef"><strong>Course Type</strong></td>
                            <td style="padding:8px 10px;border:1px solid #e8ecef">{row.get('Course Type', '')}</td>
                        </tr>
                        <tr>
                            <td style="padding:8px 10px;background:#f9fafb;border:1px solid #e8ecef"><strong>Level</strong></td>
                            <td style="padding:8px 10px;border:1px solid #e8ecef">{row.get('Level', '')}</td>
                        </tr>
                        <tr>
                            <td style="padding:8px 10px;background:#f9fafb;border:1px solid #e8ecef"><strong>Teacher</strong></td>
                            <td style="padding:8px 10px;border:1px solid #e8ecef">{row.get('Teacher', '')}</td>
                        </tr>
                        <tr>
                            <td style="padding:8px 10px;background:#f9fafb;border:1px solid #e8ecef"><strong>Number of Classes</strong></td>
                            <td style="padding:8px 10px;border:1px solid #e8ecef">{row.get('COUNT of Class No', '')}</td>
                        </tr>
                    </table>
                </td>
            </tr>

            <!-- Invoice Breakdown -->
            <tr>
                <td style="padding:0 28px 20px">
                    <p style="margin:0 0 12px;font-size:11px;text-transform:uppercase;color:#7f8c8d;letter-spacing:1px">Invoice Breakdown</p>
                    <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:14px">

                        <!-- Table Header -->
                        <tr style="background:#1a2e44;color:#fff">
                            <td style="padding:10px 12px;width:40%">Description</td>
                            <td style="padding:10px 12px;text-align:center">Classes</td>
                            <td style="padding:10px 12px;text-align:right">Rate</td>
                            <td style="padding:10px 12px;text-align:right">Amount</td>
                        </tr>

                        <!-- Line Item -->
                        <tr>
                            <td style="padding:10px 12px;border:1px solid #e0e0e0">
                                {row.get('Course_', row.get('Course', ''))} — {row.get('Course Type', '')}
                            </td>
                            <td style="padding:10px 12px;border:1px solid #e0e0e0;text-align:center">
                                {row.get('COUNT of Class No', '')}
                            </td>
                            <td style="padding:10px 12px;border:1px solid #e0e0e0;text-align:right">
                                {fmt_currency(row.get('Rate', ''))}
                            </td>
                            <td style="padding:10px 12px;border:1px solid #e0e0e0;text-align:right">
                                {fmt_currency(row.get('Amount', ''))}
                            </td>
                        </tr>

                        <!-- Subtotal (only if discounts exist) -->
                        {subtotal_row}

                        <!-- Discount Rows -->
                        {discount_rows}

                        <!-- Total -->
                        <tr style="background:#1a2e44;color:#fff">
                            <td colspan="3" style="padding:12px;font-weight:bold;font-size:15px">Total Due</td>
                            <td style="padding:12px;text-align:right;font-weight:bold;font-size:16px;color:#f0c040">
                                {fmt_currency(row.get('Amount after Discount', row.get('Amount', '')))}
                            </td>
                        </tr>

                    </table>
                </td>
            </tr>

            <!-- Footer Note -->
            <tr>
                <td style="padding:16px 28px;text-align:center;font-size:12px;color:#7f8c8d;border-top:1px solid #eee">
                    Thank you for your continued trust in New Dimension Academy.<br>
                    For any questions regarding this invoice, please reply to this email.
                </td>
            </tr>

            <!-- Footer Image -->
            {"<tr><td><img src='" + FOOTER_IMAGE_URL + "' width='620' style='display:block;width:100%'></td></tr>" if FOOTER_IMAGE_URL else ""}

        </table>

    </td></tr>
    </table>

    </body>
    </html>
    """

# ---------------- SEND EMAIL -----------------
def send_email(to_email, subject, body):
    try:
        msg = MIMEMultipart()
        msg["From"] = f"New Dimension Academy <{EMAIL_USER}>"
        msg["To"] = to_email
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "html"))

        bcc_list = ["alhuraibia@gmail.com", "dalmaznaee@gmail.com"]
        recipients = [to_email] + bcc_list
        msg["Bcc"] = ", ".join(bcc_list)

        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(EMAIL_USER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_USER, recipients, msg.as_string())

        log_message(f"✅ Invoice sent → TO: {to_email} | BCC: {', '.join(bcc_list)}")
    except Exception as e:
        log_message(f"❌ Failed to send invoice to {to_email}: {e}")

# ---------------- PROCESS INVOICES -----------------
def process_invoices():
    df = read_google_sheet()
    if df is None:
        log_message("No data to process.")
        return

    today_str = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    log_message(f"📅 Processing invoices for today: {today_str}")

    sent_count = 0
    skipped_count = 0

    for _, row in df.iterrows():
        # Normalize the Invoice Date for comparison
        invoice_date = str(row.get("Invoice Date", "")).strip().split(" ")[0]

        if invoice_date != today_str:
            skipped_count += 1
            continue

        customer_email = str(row.get("Customer Email", "")).strip()
        if not customer_email:
            log_message(f"⚠️  Skipping row — no Customer Email. Invoice #{row.get('Invoice Numeber', '?')}")
            skipped_count += 1
            continue

        student_name = row.get("Student", "")
        course = row.get("Course_", row.get("Course", ""))
        month = row.get("Service Month", "")
        year = row.get("Service Year", "")
        invoice_num = row.get("Invoice Numeber", row.get("Invoice Number", ""))

        subject = f"Invoice #{invoice_num} — {month} {year} {course} for {student_name}"

        log_message(f"📨 Sending invoice #{invoice_num} to {customer_email}")
        email_body = build_email(row)
        send_email(
            to_email=customer_email,
            subject=subject,
            body=email_body
        )
        sent_count += 1

    log_message(f"🎉 Done. Sent: {sent_count} | Skipped: {skipped_count}")

# ---------------- MAIN -----------------
if __name__ == "__main__":
    process_invoices()
