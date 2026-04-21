# -*- coding: utf-8 -*-

import os
import json
import base64
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timezone
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

try:
    from weasyprint import HTML as WeasyHTML, CSS as WeasyCSS
    WEASYPRINT_AVAILABLE = True
except ImportError:
    WEASYPRINT_AVAILABLE = False

# ---------------- CONFIGURATION -----------------
SPREADSHEET_ID = "1mhTdW15u6E-jODDpXdlJjZohVU2NHbmzF2R8TZEpIls"
RANGE_NAME = "Invoices"

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

HEADER_IMAGE_URL2 = os.getenv("HEADER_IMAGE_URL2", "")
FOOTER_IMAGE_URL  = os.getenv("FOOTER_IMAGE_URL", "")

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
    v = str(value).strip()
    return MONTH_NAMES.get(v, v)

# ---------------- CLEAN NUMERIC STRING -----------------
def clean_numeric(value) -> str:
    return str(value or '').replace(',', '').replace('$', '').replace('%', '').strip()

# ---------------- FORMAT CURRENCY -----------------
def fmt_currency(value) -> str:
    try:
        return f"{float(clean_numeric(value)):,.2f}"
    except (ValueError, TypeError):
        return "—"

# ---------------- SAFE FLOAT SUM -----------------
def safe_sum(rows, column) -> float:
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
    discount_labels = []

    for row in invoice_rows:
        for i in range(1, 4):
            d = str(row.get(f"Discount Type {i}", "")).strip()
            if d and d not in discount_labels:
                discount_labels.append(d)

    discount_amount = subtotal - total_due

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

# ---------------- BUILD INVOICE SECTIONS (shared) -----------------
def build_invoice_sections(invoice_rows):
    """Build all the inner table sections shared between email and PDF."""
    first = invoice_rows[0]

    invoice_num  = str(first.get('Invoice Numeber', first.get('Invoice Number', ''))).strip()
    invoice_date = first.get('Invoice Date', '')
    student      = first.get('Student', '')
    cust_email   = first.get('Customer Email', '')
    cust_mobile  = first.get('Customer Mobile No.', first.get('Customer Mobile No', ''))

    subtotal  = safe_sum(invoice_rows, 'Amount')
    total_due = safe_sum(invoice_rows, 'Amount after Discount')

    subtotal_fmt          = fmt_currency(subtotal)
    amount_after_discount = fmt_currency(total_due)

    has_discount = any(
        str(row.get(f"Discount Type {i}", "")).strip()
        for row in invoice_rows
        for i in range(1, 4)
    )

    course_detail_rows = build_course_details_rows(invoice_rows)
    line_item_rows     = build_line_item_rows(invoice_rows)
    discount_rows      = build_discount_rows(invoice_rows, subtotal, total_due)

    subtotal_row = f"""
        <tr style="background:#f9fafb">
            <td colspan="3" style="padding:10px 12px;border:1px solid #e0e0e0;font-weight:bold">Subtotal</td>
            <td style="padding:10px 12px;border:1px solid #e0e0e0;text-align:right;font-weight:bold">
                {subtotal_fmt}
            </td>
        </tr>""" if has_discount else ""

    return dict(
        invoice_num=invoice_num,
        invoice_date=invoice_date,
        student=student,
        cust_email=cust_email,
        cust_mobile=cust_mobile,
        amount_after_discount=amount_after_discount,
        course_detail_rows=course_detail_rows,
        line_item_rows=line_item_rows,
        subtotal_row=subtotal_row,
        discount_rows=discount_rows,
    )

# ---------------- BUILD EMAIL HTML -----------------
def build_email_html(invoice_rows, month_name: str) -> str:
    """HTML optimised for email clients (table-based layout, remote images)."""
    s = build_invoice_sections(invoice_rows)

    header_img_html = f'<tr><td><img src="{HEADER_IMAGE_URL2}" width="620" style="display:block;width:100%"></td></tr>' if HEADER_IMAGE_URL2 else ""
    footer_img_html = f'<tr><td><img src="{FOOTER_IMAGE_URL}" width="620" style="display:block;width:100%"></td></tr>' if FOOTER_IMAGE_URL else ""

    return f"""<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><title>Invoice #{s['invoice_num']}</title></head>
<body style="margin:0;padding:0;background:#f4f6f8;font-family:Arial,sans-serif">
<table width="100%" cellpadding="0" cellspacing="0">
<tr><td align="center" style="padding:30px 0">
  <table width="620" style="background:#fff;border-radius:10px;overflow:hidden;border-collapse:collapse;box-shadow:0 2px 8px rgba(0,0,0,0.08)">

    {header_img_html}

    <!-- Billed To + Invoice Meta -->
    <tr>
      <td style="padding:24px 28px;background:#f0f4f8">
        <table width="100%" cellpadding="0" cellspacing="0">
          <tr>
            <td style="vertical-align:top;width:55%">
              <p style="margin:0 0 6px;font-size:11px;text-transform:uppercase;color:#7f8c8d;letter-spacing:1px">Billed To</p>
              <p style="margin:0;font-size:18px;font-weight:bold;color:#043C4C">{s['student']}</p>
              <p style="margin:4px 0 0;font-size:13px;color:#043C4C">{s['cust_email']}</p>
              <p style="margin:2px 0 0;font-size:13px;color:#7f8c8d">{s['cust_mobile']}</p>
            </td>
            <!-- spacer -->
            <td style="width:20px"></td>
            <td align="right" style="vertical-align:top;width:45%">
              <h2 style="margin:0;font-size:22px;letter-spacing:1px;color:#043C4C">INVOICE</h2>
              <p style="margin:2px 0 0;font-size:20px;font-weight:bold;color:#f0c040">#{s['invoice_num']}</p>
              <p style="margin:8px 0 0;font-size:12px;color:#aab8c4">Date: {s['invoice_date']}</p>
            </td>
          </tr>
        </table>
      </td>
    </tr>

    <!-- Course Details -->
    <tr>
      <td style="padding:20px 28px">
        <p style="margin:0 0 12px;font-size:11px;text-transform:uppercase;color:#7f8c8d;letter-spacing:1px">Course Details</p>
        <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:14px">
          <tr style="background:#043C4C;color:#fff">
            <td style="padding:9px 10px">Course</td>
            <td style="padding:9px 10px">Type</td>
            <td style="padding:9px 10px">Level</td>
            <td style="padding:9px 10px">Teacher</td>
            <td style="padding:9px 10px;text-align:center">Classes</td>
          </tr>
          {s['course_detail_rows']}
        </table>
      </td>
    </tr>

    <!-- Invoice Breakdown -->
    <tr>
      <td style="padding:0 28px 20px">
        <p style="margin:0 0 12px;font-size:11px;text-transform:uppercase;color:#7f8c8d;letter-spacing:1px">Invoice Breakdown</p>
        <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:14px">
          <tr style="background:#043C4C;color:#fff">
            <td style="padding:10px 12px;width:40%">Description</td>
            <td style="padding:10px 12px;text-align:center">Classes</td>
            <td style="padding:10px 12px;text-align:right">Rate</td>
            <td style="padding:10px 12px;text-align:right">Amount</td>
          </tr>
          {s['line_item_rows']}
          {s['subtotal_row']}
          {s['discount_rows']}
          <tr style="background:#043C4C;color:#fff">
            <td colspan="3" style="padding:12px;font-weight:bold;font-size:15px">Total Due</td>
            <td style="padding:12px;text-align:right;font-weight:bold;font-size:16px;color:#f0c040">
              {s['amount_after_discount']}
            </td>
          </tr>
        </table>
      </td>
    </tr>

    <!-- Footer Note -->
    <tr>
      <td style="padding:16px 28px;text-align:center;font-size:12px;color:#7f8c8d;border-top:1px solid #eee">
        Thank you for your continued trust in New Dimension Academy Inc.<br>
        Please settle the due amount via e-Transfer using info@ndacademy.ca
      </td>
    </tr>

    {footer_img_html}

  </table>
</td></tr>
</table>
</body>
</html>"""

# ---------------- BUILD PDF HTML -----------------
def build_pdf_html(invoice_rows, month_name: str) -> str:
    """
    HTML optimised for WeasyPrint PDF rendering.
    - @page rule centres the content with equal margins on all sides.
    - The header section uses explicit padding-right on the left cell
      and padding-left on the right cell to guarantee visible space.
    - Fixed 620px width matches the email layout exactly.
    """
    s = build_invoice_sections(invoice_rows)

    header_img_html = f'<img src="{HEADER_IMAGE_URL2}" style="display:block;width:100%;max-width:620px">' if HEADER_IMAGE_URL2 else ""
    footer_img_html = f'<img src="{FOOTER_IMAGE_URL}" style="display:block;width:100%;max-width:620px">' if FOOTER_IMAGE_URL else ""

    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Invoice #{s['invoice_num']}</title>
<style>
  /* ── Page setup: A4, centred with equal margins ── */
  @page {{
    size: A4;
    margin: 25mm 20mm 25mm 20mm;
  }}

  * {{
    box-sizing: border-box;
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }}

  body {{
    margin: 0;
    padding: 0;
    font-family: Arial, sans-serif;
    font-size: 13px;
    color: #333;
    background: #ffffff;
    /* vertically centre the card on the page */
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    min-height: 100%;
  }}

  /* ── Outer card ── */
  .invoice-card {{
    width: 620px;
    background: #ffffff;
    border-radius: 10px;
    overflow: hidden;
    border: 1px solid #e0e0e0;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
  }}

  /* ── Header image ── */
  .header-img {{ display:block; width:100%; }}

  /* ── Billed-To / Invoice-Meta bar ── */
  .meta-bar {{
    background: #f0f4f8;
    padding: 24px 28px;
  }}
  .meta-table {{
    width: 100%;
    border-collapse: collapse;
  }}
  /* LEFT cell — billed-to */
  .meta-left {{
    vertical-align: top;
    width: 55%;
    padding-right: 40px;   /* ← the gap between the two halves */
  }}
  /* RIGHT cell — invoice number */
  .meta-right {{
    vertical-align: top;
    width: 45%;
    text-align: right;
    padding-left: 20px;    /* extra breathing room on the right */
  }}
  .label {{
    margin: 0 0 6px;
    font-size: 10px;
    text-transform: uppercase;
    color: #7f8c8d;
    letter-spacing: 1px;
  }}
  .student-name {{
    margin: 0;
    font-size: 17px;
    font-weight: bold;
    color: #043C4C;
  }}
  .meta-sub {{
    margin: 3px 0 0;
    font-size: 12px;
    color: #043C4C;
  }}
  .meta-mobile {{
    margin: 2px 0 0;
    font-size: 12px;
    color: #7f8c8d;
  }}
  .invoice-title {{
    margin: 0;
    font-size: 20px;
    letter-spacing: 1px;
    color: #043C4C;
  }}
  .invoice-num {{
    margin: 2px 0 0;
    font-size: 19px;
    font-weight: bold;
    color: #f0c040;
  }}
  .invoice-date {{
    margin: 8px 0 0;
    font-size: 11px;
    color: #aab8c4;
  }}

  /* ── Section label ── */
  .section-label {{
    margin: 0 0 12px;
    font-size: 10px;
    text-transform: uppercase;
    color: #7f8c8d;
    letter-spacing: 1px;
  }}

  /* ── Generic inner sections ── */
  .section {{ padding: 20px 28px; }}
  .section-bottom {{ padding: 0 28px 20px; }}

  /* ── Data tables ── */
  .data-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
  }}
  .data-table th {{
    background: #043C4C;
    color: #fff;
    padding: 9px 10px;
    text-align: left;
  }}
  .data-table th.right {{ text-align: right; }}
  .data-table th.center {{ text-align: center; }}
  .data-table td {{
    padding: 9px 10px;
    border: 1px solid #e8ecef;
  }}
  .data-table td.right {{ text-align: right; }}
  .data-table td.center {{ text-align: center; }}

  /* ── Total row ── */
  .total-row td {{
    background: #043C4C;
    color: #fff;
    padding: 12px;
    font-weight: bold;
    font-size: 14px;
    border: none;
  }}
  .total-amount {{
    color: #f0c040 !important;
    font-size: 15px !important;
    text-align: right;
  }}

  /* ── Subtotal row ── */
  .subtotal-row td {{
    background: #f9fafb;
    font-weight: bold;
    border: 1px solid #e0e0e0;
  }}

  /* ── Discount row ── */
  .discount-row td {{
    color: #c0392b;
    border: 1px solid #e0e0e0;
  }}

  /* ── Footer note ── */
  .footer-note {{
    padding: 16px 28px;
    text-align: center;
    font-size: 11px;
    color: #7f8c8d;
    border-top: 1px solid #eee;
  }}
</style>
</head>
<body>
<div class="invoice-card">

  {f'<div>{header_img_html}</div>' if header_img_html else ''}

  <!-- Billed To + Invoice Meta -->
  <div class="meta-bar">
    <table class="meta-table">
      <tr>
        <td class="meta-left">
          <p class="label">Billed To</p>
          <p class="student-name">{s['student']}</p>
          <p class="meta-sub">{s['cust_email']}</p>
          <p class="meta-mobile">{s['cust_mobile']}</p>
        </td>
        <td class="meta-right">
          <h2 class="invoice-title">INVOICE</h2>
          <p class="invoice-num">#{s['invoice_num']}</p>
          <p class="invoice-date">Date: {s['invoice_date']}</p>
        </td>
      </tr>
    </table>
  </div>

  <!-- Course Details -->
  <div class="section">
    <p class="section-label">Course Details</p>
    <table class="data-table">
      <tr>
        <th>Course</th>
        <th>Type</th>
        <th>Level</th>
        <th>Teacher</th>
        <th class="center">Classes</th>
      </tr>
      {s['course_detail_rows']}
    </table>
  </div>

  <!-- Invoice Breakdown -->
  <div class="section-bottom">
    <p class="section-label">Invoice Breakdown</p>
    <table class="data-table">
      <tr>
        <th style="width:40%">Description</th>
        <th class="center">Classes</th>
        <th class="right">Rate</th>
        <th class="right">Amount</th>
      </tr>
      {s['line_item_rows']}
      {s['subtotal_row']}
      {s['discount_rows']}
      <tr class="total-row">
        <td colspan="3">Total Due</td>
        <td class="total-amount">{s['amount_after_discount']}</td>
      </tr>
    </table>
  </div>

  <!-- Footer Note -->
  <div class="footer-note">
    Thank you for your continued trust in New Dimension Academy Inc.<br>
    Please settle the due amount via e-Transfer using info@ndacademy.ca
  </div>

  {f'<div>{footer_img_html}</div>' if footer_img_html else ''}

</div>
</body>
</html>"""

# ---------------- GENERATE PDF BYTES -----------------
def generate_pdf(invoice_rows, month_name: str, invoice_num: str) -> bytes | None:
    if not WEASYPRINT_AVAILABLE:
        log_message("⚠️  WeasyPrint not installed — skipping PDF attachment.")
        return None
    try:
        html = build_pdf_html(invoice_rows, month_name)
        pdf_bytes = WeasyHTML(string=html).write_pdf()
        log_message(f"✅ PDF generated for invoice #{invoice_num} ({len(pdf_bytes):,} bytes)")
        return pdf_bytes
    except Exception as e:
        log_message(f"❌ PDF generation failed for invoice #{invoice_num}: {e}")
        return None

# ---------------- SEND EMAIL -----------------
def send_email(to_email, subject, body, pdf_bytes: bytes | None = None, pdf_filename: str = "Invoice.pdf"):
    try:
        msg = MIMEMultipart()
        msg["From"]    = f"New Dimension Academy <{EMAIL_USER}>"
        msg["To"]      = to_email
        msg["Subject"] = subject

        msg.attach(MIMEText(body, "html"))

        if pdf_bytes:
            part = MIMEBase("application", "pdf")
            part.set_payload(pdf_bytes)
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", "attachment", filename=pdf_filename)
            msg.attach(part)
            log_message(f"📎 PDF attached: {pdf_filename}")

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

        # Email body and PDF are built from separate optimised templates
        email_html = build_email_html(invoice_rows, month_name)
        pdf_bytes  = generate_pdf(invoice_rows, month_name, invoice_num)
        pdf_filename = f"Invoice_{invoice_num}_{student.replace(' ', '_')}_{month_name}_{year}.pdf"

        log_message(f"📨 Sending invoice #{invoice_num} → {customer_email} ({len(invoice_rows)} line(s))")
        send_email(
            to_email     = customer_email,
            subject      = subject,
            body         = email_html,
            pdf_bytes    = pdf_bytes,
            pdf_filename = pdf_filename
        )
        sent_count += 1

    log_message(f"🎉 Done. Sent: {sent_count} | Skipped: {skipped_count}")

# ---------------- MAIN -----------------
if __name__ == "__main__":
    process_invoices()
