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
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

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

# ---------------- MONTH NUMBER TO NAME -----------------
MONTH_NAMES = {
    "01": "January",  "1": "January",
    "02": "February", "2": "February",
    "03": "March",    "3": "March",
    "04": "April",    "4": "April",
    "05": "May",      "5": "May",
    "06": "June",     "6": "June",
    "07": "July",     "7": "July",
    "08": "August",   "8": "August",
    "09": "September","9": "September",
    "10": "October",
    "11": "November",
    "12": "December",
}

def resolve_month_name(value: str) -> str:
    """Return a full month name whether the input is a number or already a name."""
    v = str(value).strip()
    return MONTH_NAMES.get(v, v)

# ---------------- CLEAN NUMERIC STRING -----------------
def clean_numeric(value) -> str:
    """Strip commas, currency symbols, percent signs, and whitespace so float() can parse it."""
    return str(value or '').replace(',', '').replace('$', '').replace('%', '').strip()

# ---------------- FORMAT CURRENCY -----------------
def fmt_currency(value) -> str:
    try:
        return f"{float(clean_numeric(value)):,.2f}"
    except (ValueError, TypeError):
        return "—"

# ---------------- SAFE FLOAT SUM -----------------
def safe_sum(rows, column) -> float:
    """Sum a column across all rows, tolerating commas, symbols, and empty cells."""
    total = 0.0
    for row in rows:
        raw = clean_numeric(row.get(column, ''))
        if raw:
            try:
                total += float(raw)
            except (ValueError, TypeError):
                pass
    return total

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

# ---------------- BUILD COURSE DETAILS ROWS -----------------
def build_course_details_rows(rows):
    """Build course detail table rows for all courses in the grouped invoice."""
    html = ""
    for i, row in enumerate(rows):
        bg          = "#f9fafb" if i % 2 == 0 else "#ffffff"
        course      = row.get('Course_', row.get('Course', ''))
        course_type = row.get('Course Type', '')
        level       = row.get('Level', '')
        teacher     = row.get('Teacher', '')
        classes     = row.get('COUNT of Class No', '')

        html += f"""
        <tr style="background:{bg}">
            <td style="padding:8px 10px;border:1px solid #e8ecef">{course}</td>
            <td style="padding:8px 10px;border:1px solid #e8ecef">{course_type}</td>
            <td style="padding:8px 10px;border:1px solid #e8ecef">{level}</td>
            <td style="padding:8px 10px;border:1px solid #e8ecef">{teacher}</td>
            <td style="padding:8px 10px;border:1px solid #e8ecef;text-align:center">{classes}</td>
        </tr>"""
    return html

# ---------------- BUILD INVOICE LINE ITEMS -----------------
def build_line_item_rows(rows):
    """Build one invoice breakdown row per course line."""
    html = ""
    for row in rows:
        course      = row.get('Course_', row.get('Course', ''))
        course_type = row.get('Course Type', '')
        classes     = row.get('COUNT of Class No', '')
        rate        = fmt_currency(row.get('Rate', ''))
        amount      = fmt_currency(row.get('Amount', ''))

        html += f"""
        <tr>
            <td style="padding:10px 12px;border:1px solid #e0e0e0">
                {course} — {course_type}
            </td>
            <td style="padding:10px 12px;border:1px solid #e0e0e0;text-align:center">{classes}</td>
            <td style="padding:10px 12px;border:1px solid #e0e0e0;text-align:right">{rate}</td>
            <td style="padding:10px 12px;border:1px solid #e0e0e0;text-align:right">{amount}</td>
        </tr>"""
    return html

# ---------------- BUILD DISCOUNT ROWS -----------------
def build_discount_rows(invoice_rows, subtotal: float, total_due: float):
    """Derive discount amount as subtotal minus total due; collect all unique discount labels."""
    discount_labels = []

    for row in invoice_rows:
        for i in range(1, 4):
            d = str(row.get(f"Discount Type {i}", "")).strip()
            if d and d not in discount_labels:
                discount_labels.append(d)

    discount_amount = subtotal - total_due

    # No labels and no meaningful discount → skip
    if not discount_labels and discount_amount <= 0.0:
        return ""

    label_text = " / ".join(discount_labels) if discount_labels else "Discount"

    return f"""
    <tr>
        <td colspan="3" style="padding:10px 12px;border:1px solid #e0e0e0;color:#c0392b">
            Discount — {label_text}
        </td>
        <td style="padding:10px 12px;border:1px solid #e0e0e0;color:#c0392b;text-align:right">
            − {fmt_currency(discount_amount)}
        </td>
    </tr>"""

# ---------------- BUILD EMAIL -----------------
def build_email(invoice_rows, month_name: str):
    first = invoice_rows[0]

    invoice_num  = str(first.get('Invoice Numeber', first.get('Invoice Number', ''))).strip()
    invoice_date = first.get('Invoice Date', '')
    student      = first.get('Student', '')
    year         = first.get('Service Year', '')
    cust_email   = first.get('Customer Email', '')
    cust_mobile  = first.get('Customer Mobile No.', first.get('Customer Mobile No', ''))

    # Subtotal = sum of Amount across all rows
    subtotal  = safe_sum(invoice_rows, 'Amount')

    # Total Due = sum of Amount after Discount across all rows
    total_due = safe_sum(invoice_rows, 'Amount after Discount')

    subtotal_fmt          = fmt_currency(subtotal)
    amount_after_discount = fmt_currency(total_due)

    # Show subtotal + discount rows only if any row carries a discount type label
    has_discount = any(
        str(row.get(f"Discount Type {i}", "")).strip()
        for row in invoice_rows
        for i in range(1, 4)
    )

    course_detail_rows = build_course_details_rows(invoice_rows)
    line_item_rows     = build_line_item_rows(invoice_rows)
    # Discount amount derived from subtotal - total_due (no dependency on Total Discount column)
    discount_rows      = build_discount_rows(invoice_rows, subtotal, total_due)

    subtotal_row = f"""
        <tr style="background:#f9fafb">
            <td colspan="3" style="padding:10px 12px;border:1px solid #e0e0e0;font-weight:bold">Subtotal</td>
            <td style="padding:10px 12px;border:1px solid #e0e0e0;text-align:right;font-weight:bold">
                {subtotal_fmt}
            </td>
        </tr>""" if has_discount else ""

    header_img_html = f'<tr><td><img src="{HEADER_IMAGE_URL}" width="620" style="display:block;width:100%"></td></tr>' if HEADER_IMAGE_URL else ""
    footer_img_html = f'<tr><td><img src="{FOOTER_IMAGE_URL}" width="620" style="display:block;width:100%"></td></tr>' if FOOTER_IMAGE_URL else ""

    return f"""
    <!DOCTYPE html>
    <html>
    <head><meta charset="UTF-8"><title>Invoice #{invoice_num}</title></head>
    <body style="margin:0;padding:0;background:#f4f6f8;font-family:Arial,sans-serif">

    <table width="100%" cellpadding="0" cellspacing="0">
    <tr><td align="center" style="padding:30px 0">

        <table width="620" style="background:#fff;border-radius:10px;overflow:hidden;border-collapse:collapse;box-shadow:0 2px 8px rgba(0,0,0,0.08)">

            {header_img_html}

            <!-- Invoice Title Bar -->
            <tr>
                <td style="padding:24px 28px;background:#1a2e44;color:#fff">
                    <table width="100%" cellpadding="0" cellspacing="0">
                        <tr>
                            <!-- Left: Academy name & address -->
                            <td style="vertical-align:top">
                                <h2 style="margin:0;font-size:22px;letter-spacing:1px">INVOICE</h2>
                                <p style="margin:8px 0 0;font-size:13px;color:#ffffff;line-height:1.7">
                                    <strong>New Dimension Academy Inc.</strong><br>
                                    Toronto, M9C 4W3 ON Canada
                                </p>
                            </td>
                            <!-- Right: Invoice meta -->
                            <td align="right" style="vertical-align:top">
                                <p style="margin:0;font-size:11px;color:#aab8c4;text-transform:uppercase;letter-spacing:1px">Invoice No.</p>
                                <p style="margin:4px 0 0;font-size:20px;font-weight:bold;color:#f0c040">
                                    #{invoice_num}
                                </p>
                                <p style="margin:8px 0 0;font-size:12px;color:#aab8c4">
                                    Date: {invoice_date}
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
                    <p style="margin:0;font-size:16px;font-weight:bold;color:#1a2e44">{student}</p>
                    <p style="margin:4px 0 0;font-size:13px;color:#555">{cust_email}</p>
                    <p style="margin:2px 0 0;font-size:13px;color:#555">{cust_mobile}</p>
                </td>
            </tr>

            <!-- Course Details -->
            <tr>
                <td style="padding:20px 28px">
                    <p style="margin:0 0 12px;font-size:11px;text-transform:uppercase;color:#7f8c8d;letter-spacing:1px">Course Details</p>
                    <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:14px">
                        <tr style="background:#1a2e44;color:#fff">
                            <td style="padding:9px 10px">Course</td>
                            <td style="padding:9px 10px">Type</td>
                            <td style="padding:9px 10px">Level</td>
                            <td style="padding:9px 10px">Teacher</td>
                            <td style="padding:9px 10px;text-align:center">Classes</td>
                        </tr>
                        {course_detail_rows}
                    </table>
                </td>
            </tr>

            <!-- Invoice Breakdown -->
            <tr>
                <td style="padding:0 28px 20px">
                    <p style="margin:0 0 12px;font-size:11px;text-transform:uppercase;color:#7f8c8d;letter-spacing:1px">Invoice Breakdown</p>
                    <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:14px">

                        <tr style="background:#1a2e44;color:#fff">
                            <td style="padding:10px 12px;width:40%">Description</td>
                            <td style="padding:10px 12px;text-align:center">Classes</td>
                            <td style="padding:10px 12px;text-align:right">Rate</td>
                            <td style="padding:10px 12px;text-align:right">Amount</td>
                        </tr>

                        {line_item_rows}
                        {subtotal_row}
                        {discount_rows}

                        <tr style="background:#1a2e44;color:#fff">
                            <td colspan="3" style="padding:12px;font-weight:bold;font-size:15px">Total Due</td>
                            <td style="padding:12px;text-align:right;font-weight:bold;font-size:16px;color:#f0c040">
                                {amount_after_discount}
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

            {footer_img_html}

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

        bcc_list   = ["alhuraibia@gmail.com", "dalmaznaee@gmail.com"]
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

    def match_today(val):
        return str(val).strip().split(" ")[0] == today_str

    df_today = df[df["Invoice Date"].apply(match_today)].copy()

    if df_today.empty:
        log_message("ℹ️  No invoices scheduled for today.")
        return

    # Handle typo in column name gracefully
    inv_col = "Invoice Numeber" if "Invoice Numeber" in df_today.columns else "Invoice Number"

    grouped = df_today.groupby(df_today[inv_col].str.strip())

    sent_count    = 0
    skipped_count = 0

    for invoice_num, group in grouped:
        invoice_rows = group.to_dict(orient="records")
        first        = invoice_rows[0]

        customer_email = str(first.get("Customer Email", "")).strip()
        if not customer_email:
            log_message(f"⚠️  Skipping invoice #{invoice_num} — no Customer Email.")
            skipped_count += 1
            continue

        student    = first.get("Student", "")
        month_raw  = first.get("Service Month", "")
        year       = first.get("Service Year", "")
        month_name = resolve_month_name(month_raw)

        subject = (
            f"Invoice #{invoice_num} | "
            f"New Dimension Academy {month_name} {year} Courses for {student}"
        )

        log_message(f"📨 Sending invoice #{invoice_num} → {customer_email} ({len(invoice_rows)} line(s))")
        email_body = build_email(invoice_rows, month_name)
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
