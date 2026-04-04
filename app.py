import os
import io
import smtplib
import pandas as pd
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from flask import Flask, request, jsonify, render_template_string

app = Flask(__name__)

GMAIL_USER = os.environ.get("GMAIL_USER", "")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")
DEFAULT_TO = os.environ.get("DEFAULT_TO", "")
DEFAULT_CC = os.environ.get("DEFAULT_CC", "")

# ── helpers ──────────────────────────────────────────────────────────────────

def parse_date(raw):
    if pd.isna(raw):
        return None
    s = str(raw).strip()
    for fmt in ("%d-%m-%Y %H:%M", "%d-%m-%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass
    # fallback: try just the date portion
    for fmt in ("%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s[:10], fmt)
        except Exception:
            pass
    return None

def fmt_date(d):
    return d.strftime("%d-%m-%Y %H:%M") if d else ""

def bucket_records(df, now):
    """Return dict of bucketed device lists."""
    d15  = now + timedelta(days=15)
    m2   = now - timedelta(days=60)
    m3   = now - timedelta(days=90)
    m4   = now - timedelta(days=120)
    m5   = now - timedelta(days=150)
    m6   = now - timedelta(days=180)

    buckets = {
        "upcoming":  [],   # expiry within next 15 days (future)
        "exp_2m":    [],   # expired within last 2 months
        "exp_3m":    [],   # expired 2-3 months ago
        "exp_4m":    [],   # expired 3-4 months ago
        "exp_5m":    [],   # expired 4-5 months ago
        "exp_6m":    [],   # expired 5-6 months ago
    }
    for _, row in df.iterrows():
        d = row["date"]
        if d is None:
            continue
        if now <= d <= d15:
            buckets["upcoming"].append(row)
        elif m2 <= d < now:
            buckets["exp_2m"].append(row)
        elif m3 <= d < m2:
            buckets["exp_3m"].append(row)
        elif m4 <= d < m3:
            buckets["exp_4m"].append(row)
        elif m5 <= d < m4:
            buckets["exp_5m"].append(row)
        elif m6 <= d < m5:
            buckets["exp_6m"].append(row)
    return buckets

def company_summary(records):
    """Return list of (company, count) sorted by company name."""
    counts = {}
    for r in records:
        counts[r["company"]] = counts.get(r["company"], 0) + 1
    items = sorted(counts.items())
    items.append(("Total", sum(c for _, c in items)))
    return items

def build_table_text(headers, rows, col_widths):
    """Build a plain-text fixed-width table."""
    lines = []
    header_line = "  ".join(h.ljust(w) for h, w in zip(headers, col_widths))
    lines.append(header_line)
    lines.append("-" * len(header_line))
    for row in rows:
        lines.append("  ".join(str(v).ljust(w) for v, w in zip(row, col_widths)))
    return "\n".join(lines)

def build_email_body(buckets, report_date_str):
    parts = []
    parts.append(f"Device Expiry Report - {report_date_str}\n")
    parts.append("=" * 55)

    # ── upcoming ──
    up = buckets["upcoming"]
    summary = company_summary(up)
    parts.append("\nUpcoming Device Expiry (Next 15 Days)\n")
    parts.append(build_table_text(
        ["Company", "Device Count"],
        summary,
        [30, 12]
    ))
    parts.append("")
    parts.append(build_table_text(
        ["Company", "Plate Number", "Installation Date"],
        [(r["company"], r["plate"], fmt_date(r["date"])) for r in up],
        [30, 16, 20]
    ))

    # ── expired buckets ──
    expired_buckets = [
        ("exp_2m", "Expired within 2 months"),
        ("exp_3m", "Expired within 3 months"),
        ("exp_4m", "Expired within 4 months"),
        ("exp_5m", "Expired within 5 months"),
        ("exp_6m", "Expired within 6 months"),
    ]

    parts.append("\n\nExpired Devices (Grouped by Months)")
    parts.append("=" * 55)

    for key, label in expired_buckets:
        recs = buckets[key]
        if not recs:
            continue
        summary = company_summary(recs)
        parts.append(f"\n{label}\n")
        parts.append(build_table_text(
            ["Company", "Device Count"],
            summary,
            [30, 12]
        ))
        parts.append("")
        parts.append(build_table_text(
            ["Company", "Plate Number", "Installation Date"],
            [(r["company"], r["plate"], fmt_date(r["date"])) for r in recs],
            [30, 16, 20]
        ))

    return "\n".join(parts)


def build_html_body(buckets, report_date_str):
    """Build a clean HTML email body matching the original format."""

    def html_summary_table(records):
        summary = company_summary(records)
        rows = "".join(
            f"<tr><td>{c}</td><td style='text-align:center'>{n}</td></tr>"
            for c, n in summary
        )
        return f"""
        <table border='1' cellpadding='5' cellspacing='0' style='border-collapse:collapse;width:100%;max-width:500px'>
          <tr style='background:#f0f0f0'><th>Company</th><th>Device Count</th></tr>
          {rows}
        </table>"""

    def html_detail_table(records):
        rows = "".join(
            f"<tr><td>{r['company']}</td><td>{r['plate']}</td><td>{fmt_date(r['date'])}</td></tr>"
            for r in records
        )
        return f"""
        <table border='1' cellpadding='5' cellspacing='0' style='border-collapse:collapse;width:100%'>
          <tr style='background:#f0f0f0'><th>Company</th><th>Plate Number</th><th>Installation Date</th></tr>
          {rows}
          <tr><td colspan='3'><b>Total: {len(records)}</b></td></tr>
        </table>"""

    expired_buckets = [
        ("exp_2m", "Expired within 2 months"),
        ("exp_3m", "Expired within 3 months"),
        ("exp_4m", "Expired within 4 months"),
        ("exp_5m", "Expired within 5 months"),
        ("exp_6m", "Expired within 6 months"),
    ]

    exp_sections = ""
    for key, label in expired_buckets:
        recs = buckets[key]
        if not recs:
            continue
        exp_sections += f"<h3>{label}</h3>{html_summary_table(recs)}<br>{html_detail_table(recs)}<br>"

    up = buckets["upcoming"]
    html = f"""
    <html><body style='font-family:Arial,sans-serif;font-size:13px'>
      <h2>Device Expiry Report - {report_date_str}</h2>
      <h3>Upcoming Device Expiry (Next 15 Days)</h3>
      {html_summary_table(up)}<br>
      {html_detail_table(up)}<br>
      <h2>Expired Devices (Grouped by Months)</h2>
      {exp_sections}
    </body></html>"""
    return html


def send_email(to, cc, subject, plain_body, html_body, attachment_bytes, attachment_name):
    msg = MIMEMultipart("mixed")
    msg["From"] = GMAIL_USER
    msg["To"] = to
    if cc:
        msg["Cc"] = cc
    msg["Subject"] = subject

    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(plain_body, "plain"))
    alt.attach(MIMEText(html_body, "html"))
    msg.attach(alt)

    part = MIMEBase("application", "octet-stream")
    part.set_payload(attachment_bytes)
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{attachment_name}"')
    msg.attach(part)

    recipients = [to] + ([cc] if cc else [])
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
        s.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        s.sendmail(GMAIL_USER, recipients, msg.as_string())


# ── routes ───────────────────────────────────────────────────────────────────

HTML_PAGE = open(os.path.join(os.path.dirname(__file__), "templates", "index.html")).read()

@app.route("/")
def index():
    return render_template_string(HTML_PAGE,
        default_to=DEFAULT_TO,
        default_cc=DEFAULT_CC,
        gmail_user=GMAIL_USER
    )

@app.route("/preview", methods=["POST"])
def preview():
    file = request.files.get("excel")
    if not file:
        return jsonify({"error": "No file uploaded"}), 400

    report_date_str = request.form.get("report_date", datetime.now().strftime("%d-%m-%Y"))
    now = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    try:
        df_raw = pd.read_excel(file, header=None)
        df_raw.columns = ["company", "plate", "raw_date"]
        df_raw["date"] = df_raw["raw_date"].apply(parse_date)
        df_raw = df_raw.dropna(subset=["date"])
        buckets = bucket_records(df_raw, now)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    plain = build_email_body(buckets, report_date_str)
    stats = {
        "total": len(df_raw),
        "upcoming": len(buckets["upcoming"]),
        "exp_2m": len(buckets["exp_2m"]),
        "exp_total": sum(len(buckets[k]) for k in ["exp_2m","exp_3m","exp_4m","exp_5m","exp_6m"])
    }
    return jsonify({"preview": plain, "stats": stats})


@app.route("/send", methods=["POST"])
def send():
    file = request.files.get("excel")
    if not file:
        return jsonify({"error": "No file uploaded"}), 400

    to_addr = request.form.get("to_email", DEFAULT_TO).strip()
    cc_addr = request.form.get("cc_email", DEFAULT_CC).strip()
    report_date_str = request.form.get("report_date", datetime.now().strftime("%d-%m-%Y"))
    now = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    file_bytes = file.read()
    filename = file.filename

    try:
        df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None)
        df_raw.columns = ["company", "plate", "raw_date"]
        df_raw["date"] = df_raw["raw_date"].apply(parse_date)
        df_raw = df_raw.dropna(subset=["date"])
        buckets = bucket_records(df_raw, now)
    except Exception as e:
        return jsonify({"error": f"Parse error: {e}"}), 500

    plain = build_email_body(buckets, report_date_str)
    html  = build_html_body(buckets, report_date_str)
    subject = f"Device Expiry Report - {report_date_str}"

    try:
        send_email(to_addr, cc_addr, subject, plain, html, file_bytes, filename)
    except Exception as e:
        return jsonify({"error": f"Email error: {e}"}), 500

    return jsonify({"success": True, "message": f"Email sent to {to_addr}"})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
