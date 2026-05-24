"""
Streamlit Community Cloud app.
Password auth via st.secrets. File upload via st.file_uploader.
Generate PPT, Word, PDF reports using Python.
"""

import io
import os
import sys
import re
import zipfile
from datetime import datetime
from pathlib import Path

import streamlit as st

HERE = Path(__file__).parent

st.set_page_config(page_title="Logistics Dashboard", page_icon="📦", layout="wide")

# ── Auth ──────────────────────────────────────────────────────────────
if "auth_ok" not in st.session_state:
    st.session_state.auth_ok = False

def login():
    st.markdown("""
    <style>
    .login-wrap{max-width:380px;margin:100px auto;padding:40px 32px;border-radius:12px;box-shadow:0 8px 32px rgba(0,0,0,.08);background:#fff;text-align:center}
    .login-wrap h1{font-size:22px;color:#1E3A5F;margin-bottom:4px}
    .login-wrap p{font-size:13px;color:#64748B;margin-bottom:28px}
    </style>
    <div class="login-wrap">
    <h1>Logistics Dashboard</h1>
    <p>Sign in with your company credentials</p>
    </div>""", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 1, 1])
    with c2:
        u = st.text_input("Username", placeholder="Enter username")
        p = st.text_input("Password", type="password", placeholder="Enter password")
        if st.button("Sign in", use_container_width=True, type="primary"):
            try:
                if u == st.secrets["auth"]["username"] and p == st.secrets["auth"]["password"]:
                    st.session_state.auth_ok = True
                    st.rerun()
                else:
                    st.error("Invalid credentials")
            except (KeyError, FileNotFoundError):
                st.error("Secrets not configured. Set [auth] username/password in Streamlit secrets.")

if not st.session_state.auth_ok:
    login()
    st.stop()

# ── Imports ───────────────────────────────────────────────────────────
sys.path.insert(0, str(HERE))
from app import parse_csv, build_report, generate_ppt, generate_docx, generate_pdf, fmt_num, fmt_money

# ── App ───────────────────────────────────────────────────────────────
st.markdown("<h1 style='color:#1E3A5F;font-size:26px;margin-bottom:4px'>Logistics Dashboard</h1>",
            unsafe_allow_html=True)
st.markdown("<p style='color:#64748B;font-size:13px;margin-bottom:16px'>"
            "Upload your customer CSV files below.</p>", unsafe_allow_html=True)

uploaded = st.file_uploader("Choose CSV files", type=["csv"], accept_multiple_files=True,
                            label_visibility="collapsed")

datasets = []
records_by_dataset = {}

if uploaded:
    for f in uploaded:
        text = f.read().decode("utf-8", errors="replace")
        records = parse_csv(text)
        if records:
            client = ""
            # Try to extract client name
            for k in records[0]:
                kl = k.lower().strip()
                if kl in ("client", "customer code", "customer"):
                    client = str(records[0].get(k, "")).replace("`", "").strip().title()
                    break
            if not client:
                client = f.name.replace(".csv", "").replace("_", " ").replace("-", " ").title()
            datasets.append({"name": client, "records": records, "file": f.name})
            records_by_dataset[f.name] = records

    if datasets:
        st.success(f"✅ {len(datasets)} dataset(s) loaded: {', '.join(d['name'] for d in datasets)}")

        # ── Dataset selector ──────────────────────────────────────────
        names = [d["name"] for d in datasets]
        if len(datasets) > 1:
            names = ["📊 All (Merged)"] + names
        sel = st.selectbox("Select dataset to view", names, label_visibility="collapsed")

        if sel == "📊 All (Merged)":
            all_recs = []
            for d in datasets:
                all_recs.extend(d["records"])
            report = build_report(all_recs)
            report["client"] = "All Customers (Consolidated)"
            label = "All Customers"
        else:
            ds = next(d for d in datasets if d["name"] == sel)
            report = build_report(ds["records"])
            report["client"] = ds["name"]
            label = ds["name"]

        # ── KPIs ──────────────────────────────────────────────────────
        k1, k2, k3, k4, k5, k6 = st.columns(6)
        with k1:
            st.metric("Shipments", fmt_num(report["shipments"]),
                      help=f"As of {datetime.now().strftime('%d %b %Y')}")
        with k2:
            st.metric("Delivered", f"{report['delivered_rate']}%",
                      f"{fmt_num(report['delivered'])} closed")
        with k3:
            st.metric("On-Time", f"{report['on_time_rate']}%",
                      f"{fmt_num(report['late_delivered'])} late")
        with k4:
            st.metric("Open Delayed", fmt_num(report['open_delayed']),
                      f"{fmt_num(report['open'])} open")
        with k5:
            st.metric("Package Value", report['value_label'],
                      f"{fmt_num(report['attempted'])} attempted")
        with k6:
            st.metric("Avg TAT", report['avg_tat_label'],
                      f"{report['weight_label']} total")

        # ── Charts ────────────────────────────────────────────────────
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("Status Breakdown")
            if report["status_entries"]:
                sc = {s: c for s, c in report["status_entries"]}
                st.bar_chart(sc, color="#2563EB")
            else:
                st.info("No status data")
        with c2:
            st.subheader("Destination States")
            if report["state_entries"]:
                state_data = {s: c for s, c in report["state_entries"]}
                st.bar_chart(state_data, color="#0F766E")
            else:
                st.info("No state data")

        c3, c4 = st.columns(2)
        with c3:
            st.subheader("Open Aging")
            if report["aging"] and any(report["aging"].values()):
                st.bar_chart(report["aging"], color="#D97706")
            else:
                st.info("No open shipments")
        with c4:
            st.subheader("Insights")
            for ins in report["insights"]:
                st.markdown(f"- {ins}")

        # ── Tables ────────────────────────────────────────────────────
        with st.expander("📋 Status Breakdown", expanded=False):
            if report["status_entries"]:
                st.dataframe({s: c for s, c in report["status_entries"]}, use_container_width=True)

        with st.expander("📋 Top Lanes", expanded=False):
            if report["lane_entries"]:
                st.dataframe({l: c for l, c in report["lane_entries"]}, use_container_width=True)

        with st.expander("📋 Priority Exceptions", expanded=False):
            now = datetime.now()
            exc = []
            for r in report["priority_records"][:20]:
                delayed = not r["is_delivered"] and r["promise"] and r["promise"] < now
                if not r["is_delivered"] or delayed:
                    exc.append({
                        "Order": r["id"][:24],
                        "Destination": f"{r['dest']}, {r['state']}",
                        "Status": "Delayed" if delayed else ("Delivered" if r["is_delivered"] else "Open"),
                        "Promise": r["promise"].strftime("%d %b") if r["promise"] else "-",
                        "Amount": fmt_money(r["amount"]),
                    })
            if exc:
                st.dataframe(exc, use_container_width=True)
            else:
                st.success("No exceptions to show")

        # ── Report generation ─────────────────────────────────────────
        st.markdown("---")
        st.subheader("📊 Generate Reports")

        # Build per-dataset reports (or merged)
        all_report_data = []
        if len(datasets) > 1:
            all_report_data.append(("All_Customers_Consolidated", report))
        # Also add individual datasets
        for d in datasets:
            if d["name"] != sel or sel == "📊 All (Merged)":
                rpt = build_report(d["records"])
                rpt["client"] = d["name"]
                all_report_data.append((d["name"], rpt))
        # Always include current view
        if not any(n == label for n, _ in all_report_data):
            all_report_data.append((label, report))

        gen_cols = st.columns(3)
        formats = [
            ("PPT", "pptx", "📊", "primary"),
            ("Word", "docx", "📝", "secondary"),
            ("PDF", "pdf", "📄", "secondary"),
        ]

        for i, (fmt_name, fmt_ext, icon, btn_type) in enumerate(formats):
            with gen_cols[i]:
                if st.button(f"{icon} Generate {fmt_name}s", use_container_width=True,
                             type="primary" if btn_type == "primary" else "secondary"):
                    with st.spinner(f"Generating {fmt_name}s..."):
                        buf = io.BytesIO()
                        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                            for cname, crpt in all_report_data:
                                crpt["client"] = cname
                                if fmt_ext == "pptx":
                                    data = generate_ppt(crpt)
                                elif fmt_ext == "docx":
                                    data = generate_docx(crpt)
                                else:
                                    data = generate_pdf(crpt)
                                safe = re.sub(r'[\\/:*?"<>|]+', '-', str(cname)).replace(' ', '_')[:50]
                                zf.writestr(f"{safe}_Logistics_Report.{fmt_ext}", data)
                        buf.seek(0)
                        st.download_button(
                            f"⬇ Download {fmt_name}s ({len(all_report_data)} files)",
                            data=buf, file_name=f"Logistics_Reports_{fmt_ext}.zip",
                            mime="application/zip", use_container_width=True,
                        )
else:
    st.info("👆 Upload one or more CSV files to get started. Supports **sonic_b2b** and **Customer MIS** formats.")

st.caption("🔒 Private — authenticated users only. Data processed in memory, not stored.")
