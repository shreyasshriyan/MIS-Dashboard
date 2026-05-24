# Logistics Dashboard

Multi-customer logistics dashboard with automated PPT, Word, and PDF report generation.
Password-protected via HTTP Basic Auth. Perfect for private company use.

---

## 🚀 Deploy on Render (FREE — 5 min setup)

### 1. Push to GitHub
```bash
# Create a NEW private repo on GitHub (don't add any files)
cd logistics-dashboard
git init
git add .
git commit -m "initial commit"
git remote add origin https://github.com/YOUR-ORG/logistics-dashboard.git
git push -u origin main
```

### 2. Deploy on Render
1. Go to https://dashboard.render.com
2. Click **"New +"** → **"Web Service"**
3. Connect your GitHub account and select the private repo
4. Fill in:
   - **Name**: `logistics-dashboard`
   - **Runtime**: `Python 3`
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120`
5. Click **"Advanced"** → **"Add Environment Variable"**
   - `RENDER_USERNAME` = `admin` (or your preferred username)
   - `RENDER_PASSWORD` = `your-strong-password`
6. Select **"Free"** plan
7. Click **"Deploy Web Service"**

### 3. Done!
Your dashboard is live at `https://logistics-dashboard.onrender.com`
Share the URL + credentials with your team.

---

## 🧪 Local Development

```bash
pip install -r requirements.txt
python3 app.py
# Open http://localhost:5000
# Login: admin / changeme
```

To change credentials locally, set env vars before starting:
```bash
export RENDER_USERNAME=myuser
export RENDER_PASSWORD=mypassword
python3 app.py
```

---

## 📁 Files

```
logistics-dashboard/
  app.py               # Flask app with built-in auth + all report generators
  index.html           # Frontend dashboard (charts, filters, KPIs)
  requirements.txt     # Python deps
  render.yaml          # Render deployment config (optional)
  sonic_b2b.csv        # Sample CSV (format 1)
  customer_mis_sample.csv  # Sample CSV (format 2)
  vendor/              # Vendor JS libs
```

---

## 📊 Usage

1. Upload one or more CSV files (drag & drop)
2. Switch between datasets or view **All (Merged)**
3. Click **PPTs**, **Word**, or **PDF** to generate per-customer reports
4. With 2+ files, a consolidated report is also created
5. All reports download as a ZIP

### Supported CSV Formats

| Format | Key columns |
|--------|-------------|
| **sonic_b2b** | `Order id`, `Client`, `Origin City`, `Destination City`, `Current Status`, `Package Amount`, `Weight`, `No of boxes` |
| **Customer MIS** | `Reference Number`, `Customer Code`, `Sender City`, `Consignee City`, `Status`, `Declared Value`, `Weight`, `Num Pieces` |

Auto-detected — no config needed.

---

## 🔒 Security

- **Built-in HTTP Basic Auth** — all routes require credentials
- Credentials set via environment variables (`RENDER_USERNAME`, `RENDER_PASSWORD`)
- Render provides **free HTTPS** (SSL) automatically
- No data stored on disk — processed entirely in memory
- GitHub repo can be **private** for extra protection
