"""
Streamlit Community Cloud app.
Host the logistics dashboard with password auth on Streamlit Cloud (free).

Setup:
  1. Push this repo to GitHub
  2. Go to https://share.streamlit.io -> deploy from your repo
  3. Set secrets: https://share.streamlit.io/dashboard -> app -> "Secrets"
     Streamlit secrets format:
       [auth]
       username = "admin"
       password = "your-secure-password"
  4. Done — your app is live with password protection.
"""

import io
import os
import sys
from pathlib import Path

import streamlit as st

HERE = Path(__file__).parent

# ── Page config ──────────────────────────────────────────────────────
st.set_page_config(
    page_title="Logistics Dashboard",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Auth check ───────────────────────────────────────────────────────
def check_auth():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if st.session_state.authenticated:
        return True

    st.markdown(
        """
        <style>
        .login-box {
            max-width: 380px; margin: 80px auto; padding: 40px 32px;
            border-radius: 12px; box-shadow: 0 8px 32px rgba(0,0,0,0.08);
            background: #fff; text-align: center;
        }
        .login-box h1 { font-size: 22px; color: #1E3A5F; margin-bottom: 4px; }
        .login-box p { font-size: 13px; color: #64748B; margin-bottom: 28px; }
        .stTextInput input { min-height: 44px; }
        </style>
        <div class="login-box">
            <h1>Logistics Dashboard</h1>
            <p>Sign in with your company credentials</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        with st.form("login_form"):
            username = st.text_input("Username", placeholder="Enter your username")
            password = st.text_input("Password", type="password", placeholder="Enter your password")
            submitted = st.form_submit_button("Sign in", use_container_width=True, type="primary")

        if submitted:
            try:
                valid_user = st.secrets["auth"]["username"]
                valid_pass = st.secrets["auth"]["password"]
            except (KeyError, FileNotFoundError):
                st.error("Secrets not configured. Set [auth] username/password in Streamlit secrets.")
                return False

            if username == valid_user and password == valid_pass:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Invalid username or password")

    return False


if not check_auth():
    st.stop()

# ── Load report generator ────────────────────────────────────────────
sys.path.insert(0, str(HERE))
from app import (
    parse_csv,
    build_report,
    generate_ppt,
    generate_docx,
    generate_pdf,
    safe_filename,
    fmt_num,
)

# ── Dashboard ────────────────────────────────────────────────────────
st.markdown(
    """
    <style>
    .block-container { padding-top: 1.5rem; padding-bottom: 1.5rem; }
    .stApp { background: #f4f7fb; }
    .uploaded { display: none; }
    </style>
    """,
    unsafe_allow_html=True,
)

# Read and inject the HTML dashboard
html_path = HERE / "index.html"
raw_html = html_path.read_text("utf-8")

# ── File upload ──────────────────────────────────────────────────────
st.markdown(
    "<h1 style='font-size:26px;color:#0F172A;margin-bottom:4px;'>"
    "Logistics Dashboard</h1>",
    unsafe_allow_html=True,
)
st.markdown(
    "<p style='color:#64748B;font-size:13px;margin-bottom:16px;'>"
    "Upload customer CSV files — supports sonic_b2b and Customer MIS formats.</p>",
    unsafe_allow_html=True,
)

uploaded_files = st.file_uploader(
    "Upload CSV files",
    type=["csv"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

st.markdown(
    "<style>div[data-testid='stFileUploader'] { margin-bottom: 8px; }</style>",
    unsafe_allow_html=True,
)

# ── Process uploaded files ───────────────────────────────────────────
report_buffers = {}
datasets_loaded = 0

if uploaded_files:
    csv_texts = {}
    for f in uploaded_files:
        try:
            text = f.read().decode("utf-8", errors="replace")
            csv_texts[f.name] = text
        except Exception as e:
            st.error(f"Could not read {f.name}: {e}")

    if csv_texts:
        datasets = []
        for fname, text in csv_texts.items():
            records = parse_csv(text)
            if records:
                client = ""
                for k in records[0]:
                    if k.lower().strip() in ("client", "customer code", "customer"):
                        client = str(records[0].get(k, "")).replace("`", "").strip().title()
                        break
                if not client:
                    client = fname.replace(".csv", "").replace("_", " ").replace("-", " ").title()
                datasets.append({"filename": fname, "records": records, "client": client})

        datasets_loaded = len(datasets)
        st.caption(f"📁 {datasets_loaded} dataset(s) loaded")

        # Generate reports in the background
        with st.spinner("Preparing reports..."):
            for ds in datasets:
                report = build_report(ds["records"])
                report["client"] = ds["client"]
                base_name = safe_filename(ds["client"])

                ppt_buf = generate_ppt(report)
                report_buffers[f"{base_name}_Logistics_Report.pptx"] = ppt_buf

                docx_buf = generate_docx(report)
                report_buffers[f"{base_name}_Logistics_Report.docx"] = docx_buf

                pdf_buf = generate_pdf(report)
                report_buffers[f"{base_name}_Logistics_Report.pdf"] = pdf_buf

            if datasets_loaded >= 2:
                all_recs = []
                for ds in datasets:
                    all_recs.extend(ds["records"])
                merged = build_report(all_recs)
                merged["client"] = "All Customers (Consolidated)"

                report_buffers["All_Customers_Consolidated_Logistics_Report.pptx"] = generate_ppt(merged)
                report_buffers["All_Customers_Consolidated_Logistics_Report.docx"] = generate_docx(merged)
                report_buffers["All_Customers_Consolidated_Logistics_Report.pdf"] = generate_pdf(merged)

# ── Download buttons ─────────────────────────────────────────────────
if report_buffers:
    mime_map = {
        "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "pdf": "application/pdf",
    }

    cols = st.columns([1, 1, 1, 3])
    labels = {"pptx": "📊 PPTs", "docx": "📝 Word", "pdf": "📄 PDFs"}
    for i, ext in enumerate(["pptx", "docx", "pdf"]):
        with cols[i]:
            matching = {k: v for k, v in report_buffers.items() if k.endswith(f".{ext}")}
            if matching:
                # Zip all matching files
                zip_buf = io.BytesIO()
                import zipfile
                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    for name, buf in matching.items():
                        zf.writestr(name, buf)
                zip_buf.seek(0)
                st.download_button(
                    label=labels[ext],
                    data=zip_buf,
                    file_name=f"Logistics_Reports_{ext}.zip",
                    mime="application/zip",
                    use_container_width=True,
                    type="primary" if ext == "pdf" else "secondary",
                )

# ── Embed full HTML dashboard ────────────────────────────────────────
dashboard_height = 1200

st.components.v1.html(raw_html, height=dashboard_height, scrolling=True)

st.caption(
    "🔒 Private — only authenticated users can access this dashboard. "
    "Data is processed in memory and not stored."
)
