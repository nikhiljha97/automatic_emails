import os
import io
import base64
import requests
import pandas as pd
from datetime import datetime, timedelta
from flask import Flask, request, jsonify, render_template_string

app = Flask(__name__)

RESEND_API_KEY = os.environ.get("RESEND_API_KEY", "")
FROM_EMAIL = os.environ.get("FROM_EMAIL", "")
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
    payload = {
        "from": FROM_EMAIL,
        "to": [to],
        "subject": subject,
        "text": plain_body,
        "html": html_body,
        "attachments": [
            {
                "filename": attachment_name,
                "content": base64.b64encode(attachment_bytes).decode("utf-8"),
            }
        ],
    }
    if cc:
        payload["cc"] = [cc]

    resp = requests.post(
        "https://api.resend.com/emails",
        headers={
            "Authorization": f"Bearer {RESEND_API_KEY}",
            "Content-Type": "application/json",
        },
        json=payload,
        timeout=30,
    )
    if resp.status_code not in (200, 201):
        raise Exception(f"Resend API error {resp.status_code}: {resp.text}")


# ── routes ───────────────────────────────────────────────────────────────────

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
  .topbar h1 { font-size: 17px; font-weight: 600; letter-spacing: 0.01em; }
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
  .status { font-size: 13px; margin-top: 10px; min-height: 20px; padding: 6px 10px; border-radius: 5px; }
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
  @media (max-width: 600px) {
    .metrics { grid-template-columns: repeat(2,1fr); }
    .form-row { grid-template-columns: 1fr; }
  }
</style>
</head>
<body>
<div class="topbar">
  <div class="dot"></div>
  <h1>Device Expiry Report — Email Sender</h1>
</div>
<div class="layout">
  <div class="card">
    <p class="card-title">1 · Upload Excel file</p>
    <div class="upload-zone" id="drop-zone" onclick="document.getElementById('xlsxFile').click()"
         ondragover="ev(event,'drag')" ondragleave="ev(event,'')" ondrop="drop(event)">
      <input type="file" id="xlsxFile" accept=".xlsx,.xls" onchange="fileChosen(this.files[0])">
      <div class="upload-icon">📂</div>
      <p><span>Click to browse</span> or drag &amp; drop</p>
      <p style="font-size:12px;color:#94a3b8;margin-top:4px">.xlsx · 3 columns: Company · Plate Number · Expiry Date</p>
    </div>
    <div id="file-pill" class="hidden">
      <div class="file-pill" id="pill-text"></div>
    </div>
  </div>
  <div class="card hidden" id="metrics-card">
    <p class="card-title">Report snapshot</p>
    <div class="metrics">
      <div class="metric"><div class="metric-label">Total devices</div><div class="metric-value" id="m-total">—</div></div>
      <div class="metric"><div class="metric-label">Expiring ≤15 days</div><div class="metric-value warn" id="m-up">—</div></div>
      <div class="metric"><div class="metric-label">Expired (2 mo)</div><div class="metric-value danger" id="m-2m">—</div></div>
      <div class="metric"><div class="metric-label">Total expired</div><div class="metric-value" id="m-exp">—</div></div>
    </div>
  </div>
  <div class="card">
    <p class="card-title">2 · Configure &amp; send</p>
    <div class="form-row">
      <div class="field">
        <label>From (Gmail)</label>
        <input type="email" id="from-addr" value="{{ gmail_user }}" readonly style="background:#f8fafc;color:#64748b">
      </div>
      <div class="field">
        <label>To</label>
        <input type="email" id="to-addr" value="{{ default_to }}" placeholder="recipient@example.com">
      </div>
    </div>
    <div class="form-row">
      <div class="field">
        <label>Report date</label>
        <input type="date" id="report-date">
      </div>
      <div class="field">
        <label>CC (optional)</label>
        <input type="email" id="cc-addr" value="{{ default_cc }}" placeholder="cc@example.com">
      </div>
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
const td = new Date();
document.getElementById('report-date').value = td.toISOString().split('T')[0];
function ev(e, cls) {
  e.preventDefault();
  document.getElementById('drop-zone').className = 'upload-zone' + (cls ? ' ' + cls : '');
}
function drop(e) {
  e.preventDefault();
  document.getElementById('drop-zone').className = 'upload-zone';
  if (e.dataTransfer.files[0]) fileChosen(e.dataTransfer.files[0]);
}
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
  el.textContent = msg;
  el.className = 'status ' + type;
  el.classList.remove('hidden');
}
function buildFormData(includeFile) {
  const fd = new FormData();
  if (includeFile) fd.append('excel', chosenFile);
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
  setBtns(true);
  setStatus('Generating preview…', 'info');
  try {
    const res = await fetch('/preview', { method: 'POST', body: buildFormData(true) });
    const data = await res.json();
    if (data.error) { setStatus('Error: ' + data.error, 'err'); setBtns(false); return; }
    document.getElementById('m-total').textContent = data.stats.total;
    document.getElementById('m-up').textContent = data.stats.upcoming;
    document.getElementById('m-2m').textContent = data.stats.exp_2m;
    document.getElementById('m-exp').textContent = data.stats.exp_total;
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
    const res = await fetch('/send', { method: 'POST', body: buildFormData(true) });
    const data = await res.json();
    if (data.error) { setStatus('Error: ' + data.error, 'err'); }
    else { setStatus('✓ ' + data.message, 'ok'); }
  } catch(e) { setStatus('Network error: ' + e.message, 'err'); }
  document.getElementById('btn-send').textContent = 'Send email';
  setBtns(false);
}
function copyPreview() {
  const txt = document.getElementById('preview-box').textContent;
  navigator.clipboard.writeText(txt).then(() => {
    const b = document.querySelector('.btn-copy');
    b.textContent = '✓ Copied!';
    setTimeout(() => b.textContent = '📋 Copy', 1800);
  });
}
</script>
</body>
</html>"""

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
