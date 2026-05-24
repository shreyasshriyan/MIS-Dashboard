# Logistics Dashboard

Multi-customer logistics dashboard with automated PPT, Word, and PDF report generation.
Supports **sonic_b2b** and **Customer MIS** CSV formats.

---

## 🚀 Quick Deploy on Streamlit Cloud (FREE — with password auth)

### 1. Push to GitHub
```bash
# Create a private GitHub repo, then:
cd logistics-dashboard
git init
git add .
git commit -m "initial commit"
git remote add origin https://github.com/YOUR-ORG/logistics-dashboard.git
git push -u origin main
```

### 2. Deploy on Streamlit Community Cloud
1. Go to https://share.streamlit.io
2. Sign in with GitHub
3. Click **"New app"** → select your private repo
4. Set:
   - **Repository**: `YOUR-ORG/logistics-dashboard`
   - **Branch**: `main`
   - **Main file**: `streamlit_app.py`
5. Click **Deploy**

### 3. Set password (Secrets)
1. After deploy, go to https://share.streamlit.io — find your app
2. Click **"⚙️ Settings"** → **"Secrets"**
3. Paste:
   ```toml
   [auth]
   username = "your-username"
   password = "your-strong-password"
   ```
4. Save — your app is now password-protected.

### 4. Share with your team
Give them the Streamlit Cloud URL. They'll need the username + password to access.

---

## 📁 File Structure

```
logistics-dashboard/
  streamlit_app.py      # ⬅ Streamlit app (deploy this on Streamlit Cloud)
  app.py                # Flask backend (also works standalone)
  index.html            # HTML dashboard with charts, filters, KPIs
  report_generator.py   # Shared Python report logic (PPT/Word/PDF)
  requirements.txt      # Python dependencies
  sonic_b2b.csv         # Sample CSV (sonic_b2b format)
  customer_mis_sample.csv  # Sample CSV (Customer MIS format)
  vendor/               # Frontend JS libraries
  .streamlit/
    secrets.toml        # Local secrets template
```

---

## Local Development

### Option A: Streamlit (recommended)
```bash
pip install -r requirements.txt
cd logistics-dashboard
streamlit run streamlit_app.py
```
Open `http://localhost:8501`. Login with credentials from `.streamlit/secrets.toml`.

### Option B: Flask (full backend)
```bash
python3 app.py
```
Open `http://localhost:5000`. No built-in auth (add nginx for production).

---

## 📊 How It Works

1. **Upload** one or more CSV files (drag & drop or click to browse)
2. **Dashboard** shows KPIs, charts, filters — switch between datasets or view merged
3. **Generate** PPT, Word, PDF reports for each customer + consolidated report
4. **Share** reports with your stakeholders

### Supported CSV Formats

| Format | Key Columns |
|--------|-------------|
| **sonic_b2b** | `Order id`, `Client`, `Origin City`, `Destination City`, `Current Status`, `Manifest Date`, `Package Amount`, `Weight`, `No of boxes` |
| **Customer MIS** | `Reference Number`, `Customer Code`, `Sender City`, `Consignee City`, `Status`, `Declared Value`, `Weight`, `Num Pieces` |

Auto-detected from column headers. No configuration needed.

---

## About Streamlit Cloud

- **Free tier**: unlimited public apps, 1 private app, 1 GB memory
- **Private GitHub repo**: works fine — Streamlit Cloud integrates with GitHub
- **Password protection**: built-in via `st.secrets` (no nginx needed)
- **HTTPS**: automatically enabled by Streamlit Cloud
- **Always-on**: your app stays running, wakes up on first visit after inactivity

---

## Security

- Credentials stored in Streamlit Secrets (encrypted at rest)
- No data written to disk — all processing in memory
- HTTPS enforced by Streamlit Cloud
- GitHub repo can be private for extra security
