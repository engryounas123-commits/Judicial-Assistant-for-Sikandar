# ⚖️ Sikandar Khan Judicial Assistant
### AI-powered legal decision-support for Pakistani courts

> This tool assists judges, lawyers, and researchers — it does NOT replace judicial authority.

---

## 🚀 Deploy in 5 Steps (Free, No Laptop Issues)

### What you will get
A real working link like `https://your-app.streamlit.app` that opens in any browser, anywhere.

---

### STEP 1 — Get a free GitHub account
Go to **github.com** → click **Sign up** → create a free account.

---

### STEP 2 — Upload these files to GitHub

1. Log in to GitHub
2. Click the **+** button (top right) → **New repository**
3. Name it: `judicial-assistant`
4. Set it to **Public**
5. Click **Create repository**
6. Now upload ALL files from this folder:
   - Click **"uploading an existing file"**
   - Drag ALL files and folders into the GitHub window
   - Click **Commit changes**

**Files to upload:**
```
app.py
requirements.txt
packages.txt
.gitignore
README.md
.streamlit/config.toml
```
> ⚠️ Do NOT upload `.streamlit/secrets.toml` — that stays private on your computer.

---

### STEP 3 — Get your Anthropic API Key (the AI brain)

1. Go to **console.anthropic.com**
2. Sign up for a free account
3. Go to **API Keys** → **Create Key**
4. Copy the key (starts with `sk-ant-...`)

---

### STEP 4 — Deploy on Streamlit Cloud (100% free)

1. Go to **share.streamlit.io**
2. Sign in with your GitHub account
3. Click **New app**
4. Select:
   - Repository: `judicial-assistant`
   - Branch: `main`
   - Main file: `app.py`
5. Click **Advanced settings** → **Secrets**
6. Paste this (replace with your real values):
   ```toml
   ANTHROPIC_API_KEY = "sk-ant-your-key-here"
   APP_USERNAME = "admin"
   APP_PASSWORD = "judicial2024"
   ```
7. Click **Save** → Click **Deploy!**

---

### STEP 5 — Share the link!

Streamlit gives you a URL like:
```
https://judicial-assistant-yourname.streamlit.app
```
Share this with anyone. Works on phone, tablet, laptop — any browser.

---

## 📋 How to Use the App

| Step | What to do |
|------|-----------|
| **Login** | Enter username/password you set in secrets |
| **Upload** | Upload case documents (PDF/DOCX/images) in Zone 1. Optionally upload reference judgment templates in Zone 2 |
| **Analyze** | Click "Run AI Case Analysis" — AI extracts parties, evidence, legal issues |
| **Judge Panel** | Click "Start Full Deliberation" — 5 AI agents deliberate (takes 1-2 mins) |
| **Judgment** | View the full formatted judgment in Word-document style |
| **Export** | Download as .docx (opens in Word/LibreOffice) or JSON |

---

## 🤖 The Five AI Judges

| Agent | Role |
|-------|------|
| Agent 1 | Pakistani Law Expert — Constitution, CPC, CrPC, Qanun-e-Shahadat |
| Agent 2 | Shariah Judge — Quran, Sunnah, Hadith, Fiqh |
| Agent 3 | Domain Specialist — property/family/tax/banking/criminal/labor/corporate |
| Agent 4 | Precedent Research — retrieves relevant Supreme Court & High Court cases |
| Agent 5 | Chief Justice — synthesizes all opinions into the final reasoned verdict |

---

## ⚠️ Important Disclaimers

- Every generated judgment carries an explicit AI disclaimer
- Unverified case citations are clearly marked **UNVERIFIED REFERENCE**
- This tool is for decision-support only — all outputs must be reviewed by a human judicial officer
- Do not use for real court filings without independent legal verification

---

## 💰 Cost

| Component | Cost |
|-----------|------|
| GitHub | Free |
| Streamlit Cloud | Free |
| Anthropic API | ~$0.01–0.10 per case analysis (pay as you go) |

---

## 🆘 Troubleshooting

**"App not loading"** → Wait 2-3 minutes after deploy; first load takes time.

**"API Key error"** → Double-check secrets in Streamlit dashboard → App Settings → Secrets.

**"File upload error"** → Supported formats: PDF, DOCX, JPG, PNG, TIFF.

**"Analysis returned error"** → Your Anthropic API key may be invalid or have no credits.
