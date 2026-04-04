import os
import io
import base64
import calendar
import requests
import pandas as pd
from datetime import datetime, timedelta
from flask import Flask, request, jsonify, render_template_string

app = Flask(__name__)

RESEND_API_KEY = os.environ.get("RESEND_API_KEY", "")
FROM_EMAIL     = os.environ.get("FROM_EMAIL", "")
DEFAULT_TO     = os.environ.get("DEFAULT_TO", "")
DEFAULT_CC     = os.environ.get("DEFAULT_CC", "")

# ── date helpers ──────────────────────────────────────────────────────────────

def parse_date(raw):
    if pd.isna(raw):
        return None
    s = str(raw).strip()
    if s in ("--", "", "nan", "None"):
        return None
    for fmt in ("%d-%m-%Y %H:%M", "%d-%m-%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass
    for fmt in ("%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s[:10], fmt)
        except Exception:
            pass
    return None

def fmt_date(d):
    return d.strftime("%d-%m-%Y %H:%M") if d else ""

def get_anchor(upload_date):
    """
    Day 1–15  of month → anchor = 1st of that month
    Day 16–31 of month → anchor = 16th of that month
    """
    day = upload_date.day
    if day <= 15:
        return upload_date.replace(day=1,  hour=0, minute=0, second=0, microsecond=0)
    else:
        return upload_date.replace(day=16, hour=0, minute=0, second=0, microsecond=0)

def get_upcoming_end(anchor):
    """
    Anchor = 1st  → end = 15th of same month
    Anchor = 16th → end = last day of same month
    """
    if anchor.day == 1:
        return anchor.replace(day=15)
    else:
        last_day = calendar.monthrange(anchor.year, anchor.month)[1]
        return anchor.replace(day=last_day)

# ── Excel loader ──────────────────────────────────────────────────────────────

def find_header_row(raw_df):
    """Scan rows to find the one containing 'Installation Date'."""
    for i, row in raw_df.iterrows():
        if any(str(v).strip() == "Installation Date" for v in row.values):
            return i
    return None

def load_dataframe(file_obj):
    """
    Read Excel with auto-detected header row.
    Finds 'Company' and 'Plate Number' and 'Installation Date' by column name.
    Returns clean df with columns: company, plate, date.
    """
    raw = pd.read_excel(file_obj, header=None)
    header_row = find_header_row(raw)

    if header_row is None:
        raise ValueError("Could not find a column named 'Installation Date' in this file.")

    # Re-read using the correct header row
    file_obj.seek(0)
    df = pd.read_excel(file_obj, header=header_row)

    # Strip whitespace from all column names
    df.columns = [str(c).strip() for c in df.columns]

    # Find required columns (case-insensitive)
    col_map = {c.lower(): c for c in df.columns}

    def get_col(name):
        key = name.lower()
        if key not in col_map:
            raise ValueError(f"Column '{name}' not found. Available: {list(df.columns)}")
        return col_map[key]

    company_col  = get_col("Company")
    plate_col    = get_col("Plate Number")
    date_col     = get_col("Installation Date")

    out = pd.DataFrame()
    out["company"] = df[company_col].astype(str).str.strip()
    out["plate"]   = df[plate_col].astype(str).str.strip()
    out["date"]    = df[date_col].apply(parse_date)

    # Drop rows where date couldn't be parsed or company is blank/nan
    out = out[out["company"].notna() & (out["company"] != "nan") & (out["company"] != "")]
    out = out.dropna(subset=["date"])
    return out

# ── bucketing ─────────────────────────────────────────────────────────────────

def bucket_records(df, upload_date):
    """
    Anchor = 1st if upload day ≤ 15, else 16th.
    Upcoming  : anchor → upcoming_end  (forward window within month half)
    exp_2m    : anchor-60d  → anchor
    exp_3m    : anchor-90d  → anchor-60d
    exp_4m    : anchor-120d → anchor-90d
    exp_5m    : anchor-150d → anchor-120d
    exp_6m    : anchor-180d → anchor-150d
    """
    anchor       = get_anchor(upload_date)
    upcoming_end = get_upcoming_end(anchor)
    m2 = anchor - timedelta(days=60)
    m3 = anchor - timedelta(days=90)
    m4 = anchor - timedelta(days=120)
    m5 = anchor - timedelta(days=150)
    m6 = anchor - timedelta(days=180)

    buckets = {
        "upcoming":     [],
        "exp_2m":       [],
        "exp_3m":       [],
        "exp_4m":       [],
        "exp_5m":       [],
        "exp_6m":       [],
        "_anchor":      anchor,
        "_upcoming_end":upcoming_end,
    }

    for _, row in df.iterrows():
        d = row["date"]
        if d is None:
            continue
        if anchor <= d <= upcoming_end:
            buckets["upcoming"].append(row)
        elif m2 <= d < anchor:
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

# ── report builders ───────────────────────────────────────────────────────────

def company_summary(records):
    counts = {}
    for r in records:
        counts[r["company"]] = counts.get(r["company"], 0) + 1
    items = sorted(counts.items())
    items.append(("Total", sum(c for _, c in items)))
    return items

def build_table_text(headers, rows, col_widths):
    lines = []
    header_line = "  ".join(h.ljust(w) for h, w in zip(headers, col_widths))
    lines.append(header_line)
    lines.append("-" * len(header_line))
    for row in rows:
        lines.append("  ".join(str(v).ljust(w) for v, w in zip(row, col_widths)))
    return "\n".join(lines)

def build_email_body(buckets, report_date_str):
    anchor       = buckets["_anchor"]
    upcoming_end = buckets["_upcoming_end"]
    parts = []
    parts.append(f"Device Expiry Report - {report_date_str}\n")
    parts.append(f"Anchor date: {anchor.strftime('%d-%m-%Y')}  |  "
                 f"Upcoming window: {anchor.strftime('%d-%m-%Y')} to {upcoming_end.strftime('%d-%m-%Y')}")
    parts.append("=" * 60)

    up = buckets["upcoming"]
    parts.append(f"\nUpcoming Device Expiry ({anchor.strftime('%d-%m-%Y')} – {upcoming_end.strftime('%d-%m-%Y')})\n")
    parts.append(build_table_text(["Company", "Device Count"], company_summary(up), [32, 12]))
    parts.append("")
    parts.append(build_table_text(
        ["Company", "Plate Number", "Installation Date"],
        [(r["company"], r["plate"], fmt_date(r["date"])) for r in up],
        [32, 18, 20]
    ))

    expired_buckets = [
        ("exp_2m", "Expired within 2 months"),
        ("exp_3m", "Expired within 3 months"),
        ("exp_4m", "Expired within 4 months"),
        ("exp_5m", "Expired within 5 months"),
        ("exp_6m", "Expired within 6 months"),
    ]

    parts.append("\n\nExpired Devices (Grouped by Months)")
    parts.append("=" * 60)

    for key, label in expired_buckets:
        recs = buckets[key]
        if not recs:
            continue
        parts.append(f"\n{label}\n")
        parts.append(build_table_text(["Company", "Device Count"], company_summary(recs), [32, 12]))
        parts.append("")
        parts.append(build_table_text(
            ["Company", "Plate Number", "Installation Date"],
            [(r["company"], r["plate"], fmt_date(r["date"])) for r in recs],
            [32, 18, 20]
        ))

    return "\n".join(parts)

def build_html_body(buckets, report_date_str):
    anchor       = buckets["_anchor"]
    upcoming_end = buckets["_upcoming_end"]

    def html_summary_table(records):
        summary = company_summary(records)
        rows = "".join(
            f"<tr><td>{c}</td><td style='text-align:center'>{n}</td></tr>"
            for c, n in summary
        )
        return f"""<table border='1' cellpadding='5' cellspacing='0'
            style='border-collapse:collapse;width:100%;max-width:520px'>
          <tr style='background:#f0f0f0'><th>Company</th><th>Device Count</th></tr>
          {rows}</table>"""

    def html_detail_table(records):
        rows = "".join(
            f"<tr><td>{r['company']}</td><td>{r['plate']}</td><td>{fmt_date(r['date'])}</td></tr>"
            for r in records
        )
        return f"""<table border='1' cellpadding='5' cellspacing='0'
            style='border-collapse:collapse;width:100%'>
          <tr style='background:#f0f0f0'>
            <th>Company</th><th>Plate Number</th><th>Installation Date</th></tr>
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
    return f"""<html><body style='font-family:Arial,sans-serif;font-size:13px'>
      <h2>Device Expiry Report - {report_date_str}</h2>
      <p style='color:#555'>Anchor: {anchor.strftime('%d-%m-%Y')} &nbsp;|&nbsp;
         Upcoming window: {anchor.strftime('%d-%m-%Y')} – {upcoming_end.strftime('%d-%m-%Y')}</p>
      <h3>Upcoming Device Expiry ({anchor.strftime('%d-%m-%Y')} – {upcoming_end.strftime('%d-%m-%Y')})</h3>
      {html_summary_table(up)}<br>{html_detail_table(up)}<br>
      <h2>Expired Devices (Grouped by Months)</h2>
      {exp_sections}
    </body></html>"""

# ── email sender ──────────────────────────────────────────────────────────────

def send_email(to, cc, subject, plain_body, html_body, attachment_bytes, attachment_name):
    payload = {
        "from": FROM_EMAIL,
        "to":   [to],
        "subject": subject,
        "text": plain_body,
        "html": html_body,
        "attachments": [{
            "filename": attachment_name,
            "content":  base64.b64encode(attachment_bytes).decode("utf-8"),
        }],
    }
    if cc:
        payload["cc"] = [cc]

    resp = requests.post(
        "https://api.resend.com/emails",
        headers={
            "Authorization": f"Bearer {RESEND_API_KEY}",
            "Content-Type":  "application/json",
        },
        json=payload,
        timeout=30,
    )
    if resp.status_code not in (200, 201):
        raise Exception(f"Resend API error {resp.status_code}: {resp.text}")

# ── HTML page ─────────────────────────────────────────────────────────────────

HTML_PAGE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Device Expiry Report</title>
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: Arial, sans-serif; font-size: 14px; background: #f4f6f9; color: #222; min-height: 100vh; }
  .topbar { background: #1a1a2e; color: #fff; padding: 14px 28px; display: flex; align-items: center; gap: 12px; }
  .topbar h1 { font-size: 17px; font-weight: 600; }
  .topbar .dot { width: 10px; height: 10px; border-radius: 50%; background: #4ade80; }
  .layout { max-width: 860px; margin: 0 auto; padding: 28px 20px; display: grid; gap: 18px; }
  .card { background: #fff; border-radius: 10px; border: 1px solid #e2e8f0; padding: 22px 24px; }
  .card-title { font-size: 13px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.06em; color: #64748b; margin-bottom: 16px; }
  .upload-zone { border: 2px dashed #cbd5e1; border-radius: 8px; padding: 32px 20px; text-align: center; cursor: pointer; transition: border-color 0.2s, background 0.2s; }
  .upload-zone:hover, .upload-zone.drag { border-color: #3b82f6; background: #eff6ff; }
  .upload-zone input { display: none; }
  .upload-icon { font-size: 36px; margin-bottom: 8px; }
  .upload-zone p { color: #475569; font-size: 14px; }
  .upload-zone span { color: #3b82f6; font-weight: 600; }
  .file-pill { display: inline-flex; align-items: center; gap: 8px; background: #f0fdf4; border: 1px solid #86efac; border-radius: 20px; padding: 4px 12px; font-size: 13px; color: #166534; margin-top: 10px; }
  .metrics { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; }
  .metric { background: #f8fafc; border-radius: 8px; padding: 14px 16px; border: 1px solid #e2e8f0; }
  .metric-label { font-size: 11px; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 6px; }
  .metric-value { font-size: 26px; font-weight: 700; color: #1e293b; }
  .metric-value.warn { color: #d97706; }
  .metric-value.danger { color: #dc2626; }
  .anchor-info { font-size: 12px; color: #64748b; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 8px 12px; margin-top: 12px; }
  .form-row { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-bottom: 14px; }
  .field label { display: block; font-size: 12px; color: #64748b; font-weight: 600; margin-bottom: 5px; }
  .field input { width: 100%; padding: 9px 12px; border: 1px solid #cbd5e1; border-radius: 6px; font-size: 14px; color: #1e293b; background: #fff; }
  .field input:focus { outline: none; border-color: #3b82f6; box-shadow: 0 0 0 3px rgba(59,130,246,0.1); }
  .btn-row { display: flex; gap: 10px; margin-top: 4px; }
  .btn { padding: 10px 22px; border-radius: 7px; font-size: 14px; font-weight: 600; cursor: pointer; border: none; transition: opacity 0.15s, transform 0.1s; }
  .btn:active { transform: scale(0.98); }
  .btn:disabled { opacity: 0.4; cursor: not-allowed; }
  .btn-preview { background: #1e293b; color: #fff; }
  .btn-send { background: #16a34a; color: #fff; }
  .btn-preview:hover:not(:disabled) { background: #0f172a; }
  .btn-send:hover:not(:disabled) { background: #15803d; }
  .status { font-size: 13px; margin-top: 10px; padding: 6px 10px; border-radius: 5px; }
  .status.info { background: #eff6ff; color: #1d4ed8; }
  .status.ok { background: #f0fdf4; color: #166534; }
  .status.err { background: #fef2f2; color: #991b1b; }
  .preview-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; }
  .btn-copy { padding: 6px 14px; border: 1px solid #cbd5e1; border-radius: 6px; background: #fff; font-size: 12px; cursor: pointer; font-weight: 600; color: #475569; }
  .btn-copy:hover { background: #f8fafc; }
  #preview-box { font-family: 'Courier New', monospace; font-size: 12px; line-height: 1.65; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 16px; white-space: pre; overflow-x: auto; max-height: 440px; overflow-y: auto; color: #1e293b; }
  .spinner { display: inline-block; width: 14px; height: 14px; border: 2px solid rgba(255,255,255,0.4); border-top-color: #fff; border-radius: 50%; animation: spin 0.6s linear infinite; vertical-align: middle; margin-right: 6px; }
  @keyframes spin { to { transform: rotate(360deg); } }
  .hidden { display: none !important; }
  @media (max-width: 600px) { .metrics { grid-template-columns: repeat(2,1fr); } .form-row { grid-template-columns: 1fr; } }
</style>
</head>
<body>
<div class="topbar"><div class="dot"></div><h1>Device Expiry Report — Email Sender</h1></div>
<div class="layout">
  <div class="card">
    <p class="card-title">1 · Upload Excel file</p>
    <div class="upload-zone" id="drop-zone" onclick="document.getElementById('xlsxFile').click()"
         ondragover="ev(event,'drag')" ondragleave="ev(event,'')" ondrop="drop(event)">
      <input type="file" id="xlsxFile" accept=".xlsx,.xls" onchange="fileChosen(this.files[0])">
      <div class="upload-icon">📂</div>
      <p><span>Click to browse</span> or drag &amp; drop</p>
      <p style="font-size:12px;color:#94a3b8;margin-top:4px">Must contain columns: Company · Plate Number · Installation Date</p>
    </div>
    <div id="file-pill" class="hidden"><div class="file-pill" id="pill-text"></div></div>
  </div>

  <div class="card hidden" id="metrics-card">
    <p class="card-title">Report snapshot</p>
    <div class="metrics">
      <div class="metric"><div class="metric-label">Total devices</div><div class="metric-value" id="m-total">—</div></div>
      <div class="metric"><div class="metric-label">Upcoming</div><div class="metric-value warn" id="m-up">—</div></div>
      <div class="metric"><div class="metric-label">Expired (2 mo)</div><div class="metric-value danger" id="m-2m">—</div></div>
      <div class="metric"><div class="metric-label">Total expired</div><div class="metric-value" id="m-exp">—</div></div>
    </div>
    <div class="anchor-info" id="anchor-info"></div>
  </div>

  <div class="card">
    <p class="card-title">2 · Configure &amp; send</p>
    <div class="form-row">
      <div class="field"><label>From (Gmail)</label>
        <input type="email" id="from-addr" value="{{ gmail_user }}" readonly style="background:#f8fafc;color:#64748b"></div>
      <div class="field"><label>To</label>
        <input type="email" id="to-addr" value="{{ default_to }}" placeholder="recipient@example.com"></div>
    </div>
    <div class="form-row">
      <div class="field"><label>Report date</label>
        <input type="date" id="report-date"></div>
      <div class="field"><label>CC (optional)</label>
        <input type="email" id="cc-addr" value="{{ default_cc }}" placeholder="cc@example.com"></div>
    </div>
    <div class="btn-row">
      <button class="btn btn-preview" id="btn-preview" onclick="doPreview()" disabled>Preview email</button>
      <button class="btn btn-send" id="btn-send" onclick="doSend()" disabled>Send email</button>
    </div>
    <div class="status hidden" id="status-bar"></div>
  </div>

  <div class="card hidden" id="preview-card">
    <div class="preview-header">
      <p class="card-title" style="margin:0">3 · Email preview</p>
      <button class="btn-copy" onclick="copyPreview()">📋 Copy</button>
    </div>
    <div id="preview-box"></div>
  </div>
</div>

<script>
let chosenFile = null;
document.getElementById('report-date').value = new Date().toISOString().split('T')[0];

function ev(e, cls) { e.preventDefault(); document.getElementById('drop-zone').className = 'upload-zone' + (cls ? ' '+cls : ''); }
function drop(e) { e.preventDefault(); document.getElementById('drop-zone').className = 'upload-zone'; if (e.dataTransfer.files[0]) fileChosen(e.dataTransfer.files[0]); }
function fileChosen(f) {
  if (!f) return;
  chosenFile = f;
  document.getElementById('pill-text').textContent = '✓ ' + f.name + '  (' + (f.size/1024).toFixed(0) + ' KB)';
  document.getElementById('file-pill').classList.remove('hidden');
  document.getElementById('btn-preview').disabled = false;
  document.getElementById('btn-send').disabled = false;
  document.getElementById('metrics-card').classList.add('hidden');
  document.getElementById('preview-card').classList.add('hidden');
  setStatus('', '');
}
function setStatus(msg, type) {
  const el = document.getElementById('status-bar');
  if (!msg) { el.classList.add('hidden'); return; }
  el.textContent = msg; el.className = 'status ' + type; el.classList.remove('hidden');
}
function buildFormData() {
  const fd = new FormData();
  fd.append('excel', chosenFile);
  fd.append('to_email', document.getElementById('to-addr').value);
  fd.append('cc_email', document.getElementById('cc-addr').value);
  const d = document.getElementById('report-date').value;
  const [y,m,day] = d.split('-');
  fd.append('report_date', `${day}-${m}-${y}`);
  return fd;
}
function setBtns(loading) {
  document.getElementById('btn-preview').disabled = loading;
  document.getElementById('btn-send').disabled = loading;
}
async function doPreview() {
  if (!chosenFile) return;
  setBtns(true); setStatus('Generating preview…', 'info');
  try {
    const res = await fetch('/preview', { method: 'POST', body: buildFormData() });
    const data = await res.json();
    if (data.error) { setStatus('Error: ' + data.error, 'err'); setBtns(false); return; }
    document.getElementById('m-total').textContent = data.stats.total;
    document.getElementById('m-up').textContent = data.stats.upcoming;
    document.getElementById('m-2m').textContent = data.stats.exp_2m;
    document.getElementById('m-exp').textContent = data.stats.exp_total;
    document.getElementById('anchor-info').textContent =
      'Anchor: ' + data.stats.anchor + '  |  Upcoming window: ' + data.stats.upcoming_window;
    document.getElementById('metrics-card').classList.remove('hidden');
    document.getElementById('preview-box').textContent = data.preview;
    document.getElementById('preview-card').classList.remove('hidden');
    setStatus('Preview ready. Review below, then click Send.', 'ok');
  } catch(e) { setStatus('Network error: ' + e.message, 'err'); }
  setBtns(false);
}
async function doSend() {
  if (!chosenFile) return;
  const to = document.getElementById('to-addr').value.trim();
  if (!to) { setStatus('Please enter a recipient email address.', 'err'); return; }
  if (!confirm(`Send email to ${to}?`)) return;
  setBtns(true);
  document.getElementById('btn-send').innerHTML = '<span class="spinner"></span>Sending…';
  setStatus('Sending email with attachment…', 'info');
  try {
    const res = await fetch('/send', { method: 'POST', body: buildFormData() });
    const data = await res.json();
    if (data.error) { setStatus('Error: ' + data.error, 'err'); }
    else { setStatus('✓ ' + data.message, 'ok'); }
  } catch(e) { setStatus('Network error: ' + e.message, 'err'); }
  document.getElementById('btn-send').textContent = 'Send email';
  setBtns(false);
}
function copyPreview() {
  navigator.clipboard.writeText(document.getElementById('preview-box').textContent).then(() => {
    const b = document.querySelector('.btn-copy');
    b.textContent = '✓ Copied!';
    setTimeout(() => b.textContent = '📋 Copy', 1800);
  });
}
</script>
</body>
</html>"""

# ── routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template_string(HTML_PAGE,
        default_to=DEFAULT_TO,
        default_cc=DEFAULT_CC,
        gmail_user=FROM_EMAIL
    )

@app.route("/preview", methods=["POST"])
def preview():
    file = request.files.get("excel")
    if not file:
        return jsonify({"error": "No file uploaded"}), 400

    report_date_str = request.form.get("report_date", datetime.now().strftime("%d-%m-%Y"))
    upload_date = datetime.now()

    try:
        df = load_dataframe(file)
        buckets = bucket_records(df, upload_date)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    anchor       = buckets["_anchor"]
    upcoming_end = buckets["_upcoming_end"]
    plain = build_email_body(buckets, report_date_str)
    stats = {
        "total":           len(df),
        "upcoming":        len(buckets["upcoming"]),
        "exp_2m":          len(buckets["exp_2m"]),
        "exp_total":       sum(len(buckets[k]) for k in ["exp_2m","exp_3m","exp_4m","exp_5m","exp_6m"]),
        "anchor":          anchor.strftime("%d-%m-%Y"),
        "upcoming_window": f"{anchor.strftime('%d-%m-%Y')} → {upcoming_end.strftime('%d-%m-%Y')}",
    }
    return jsonify({"preview": plain, "stats": stats})


@app.route("/send", methods=["POST"])
def send():
    file = request.files.get("excel")
    if not file:
        return jsonify({"error": "No file uploaded"}), 400

    to_addr         = request.form.get("to_email", DEFAULT_TO).strip()
    cc_addr         = request.form.get("cc_email",  DEFAULT_CC).strip()
    report_date_str = request.form.get("report_date", datetime.now().strftime("%d-%m-%Y"))
    upload_date     = datetime.now()

    file_bytes = file.read()
    filename   = file.filename

    try:
        df      = load_dataframe(io.BytesIO(file_bytes))
        buckets = bucket_records(df, upload_date)
    except Exception as e:
        return jsonify({"error": f"Parse error: {e}"}), 500

    plain   = build_email_body(buckets, report_date_str)
    html    = build_html_body(buckets, report_date_str)
    subject = f"Device Expiry Report - {report_date_str}"

    try:
        send_email(to_addr, cc_addr, subject, plain, html, file_bytes, filename)
    except Exception as e:
        return jsonify({"error": f"Email error: {e}"}), 500

    return jsonify({"success": True, "message": f"Email sent to {to_addr}"})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
