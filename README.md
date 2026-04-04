# Device Expiry Report — Email Sender

A web app that lets your dad upload an Excel file and automatically send
a formatted Device Expiry Report email (with the Excel attached) — from any browser.

---

## How it works

1. Dad opens the app URL in any browser
2. Uploads the Excel file (3 columns: Company | Plate Number | Expiry Date)
3. Clicks **Preview email** to review
4. Clicks **Send email** — done. Email arrives with the Excel attached.

---

## One-time setup (you do this once, your dad never has to)

### Step 1 — Get a Gmail App Password

Gmail requires an "App Password" for server-side sending.

1. Go to your Google Account → **Security**
2. Make sure **2-Step Verification** is ON
3. Search for **"App Passwords"** in the search bar at the top
4. Create a new App Password → name it "Device Expiry App"
5. Copy the 16-character password (looks like: `abcd efgh ijkl mnop`)
6. Remove the spaces when you use it → `abcdefghijklmnop`

---

### Step 2 — Deploy to Railway (free, recommended)

**Railway gives you a permanent URL in ~3 minutes.**

1. Create a free account at https://railway.app
2. Click **"New Project"** → **"Deploy from GitHub repo"**
   - Or use **"Deploy from local"** if you don't want to use GitHub
3. Upload / push this entire folder
4. Once deployed, go to your project → **Variables** tab
5. Add these environment variables:

   | Variable | Value |
   |---|---|
   | `GMAIL_USER` | `rohitpathak88@gmail.com` |
   | `GMAIL_APP_PASSWORD` | your 16-char app password (no spaces) |
   | `DEFAULT_TO` | `navinj@fnsnigeria.com` |
   | `DEFAULT_CC` | *(leave blank or add a CC address)* |

6. Railway auto-detects the `Procfile` and starts the app
7. Go to **Settings → Networking → Generate Domain** to get your public URL

That's it — share the URL with your dad.

---

### Alternative: Deploy to Render (also free)

1. Create account at https://render.com
2. New → **Web Service** → connect your GitHub repo (or upload files)
3. Set:
   - **Build command**: `pip install -r requirements.txt`
   - **Start command**: `gunicorn app:app`
4. Add the same environment variables under **Environment**
5. Click **Deploy** → copy the `.onrender.com` URL

---

### Step 3 — Test it

1. Open your URL
2. Upload the sample Excel
3. Click Preview → verify the report looks correct
4. Click Send → check the inbox

---

## Local testing (optional)

```bash
# Install dependencies
pip install -r requirements.txt

# Set environment variables
cp .env.example .env
# Edit .env and fill in your values

# Run
python app.py
# Open http://localhost:5000
```

---

## Excel format required

The Excel file must have **no header row**, just 3 columns:

| Column A | Column B | Column C |
|---|---|---|
| Company name | Plate number | Expiry date |
| AANDG TIPPERS | AAB312XD | 04-04-2026 00:00 |

Accepted date formats: `DD-MM-YYYY HH:MM` or `DD-MM-YYYY` or `YYYY-MM-DD`

---

## What the email contains

- **Subject**: `Device Expiry Report - DD-MM-YYYY`
- **Upcoming (next 15 days)**: company summary table + full device list
- **Expired within 2 months**: company summary + device list
- **Expired within 3 months**: company summary + device list
- **Expired within 4 months**: company summary + device list
- **Expired within 5 months**: company summary + device list
- **Expired within 6 months**: company summary + device list
- **Attachment**: the original Excel file

---

## Files in this project

```
device-expiry-app/
├── app.py              ← main server logic
├── templates/
│   └── index.html      ← the web UI your dad uses
├── requirements.txt    ← Python dependencies
├── Procfile            ← tells Railway/Render how to start the app
├── .env.example        ← template for environment variables
├── .gitignore
└── README.md
```
