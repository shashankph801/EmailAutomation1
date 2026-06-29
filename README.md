# Resume Email Agent

A fully independent resume email agent — no AI API, no paid services.
Uses `pypdf` + regex to extract emails, and sends mail via the **Gmail API**.

---

## Setup

### 1. Install dependencies
```bash
pip install -r requirements.txt
```

### 2. Configure Gmail API credentials

#### A. Google Cloud Console (one-time)

1. Create a project at https://console.cloud.google.com/
2. Enable **Gmail API** (APIs & Services → Library)
3. Configure **OAuth consent screen** (External is fine for personal Gmail; add your email as a test user)
4. Create **OAuth 2.0 Client ID** → Application type: **Desktop app**
5. Download the client JSON file

#### B. Get a refresh token (one-time, run on your Mac)

```bash
source venv/bin/activate
python get_gmail_refresh_token.py --credentials /path/to/client_secret.json
```

Sign in with the Gmail account that will send emails. The script prints env vars to copy.

#### C. Add to `.env` (local) or Render environment variables

```env
SENDER_EMAIL=you@gmail.com
GOOGLE_CLIENT_ID=...
GOOGLE_CLIENT_SECRET=...
GOOGLE_REFRESH_TOKEN=...
```

`SENDER_EMAIL` must match the Google account used to generate the refresh token.

### 3. Run the app
```bash
python app.py
```

### 4. Open in browser
```
http://localhost:5000
```

---

## How it works

1. **Upload resume** (PDF or DOCX)
2. `pypdf` / `python-docx` extracts raw text from the file
3. Regex finds the email address in the text
4. You attach your JD + Assessment files
5. Email is sent via Gmail API (HTTPS — works on Render and locally)

---

## Stack
- **Backend**: Flask (Python)
- **PDF parsing**: pypdf
- **DOCX parsing**: python-docx
- **Email extraction**: regex
- **Email sending**: Gmail API
- **Frontend**: Vanilla HTML/CSS/JS

Zero AI APIs. Zero paid services.
