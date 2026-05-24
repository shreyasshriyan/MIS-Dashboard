#!/usr/bin/env python3
"""
Logistics Dashboard — Flask backend for server-side report generation.
Serves the static dashboard and provides API endpoints for PPT, Word, and PDF generation.
"""

import io
import os
import re
import zipfile
from datetime import datetime, date
from pathlib import Path

from flask import Flask, jsonify, request, send_file, render_template_string
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from docx import Document
from docx.shared import Inches as DocInches, Pt as DocPt, RGBColor as DocRGB, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from fpdf import FPDF

# ── App setup ──────────────────────────────────────────────────────────

HERE = Path(__file__).parent
app = Flask(__name__, static_folder=str(HERE), static_url_path='')

# ── Colour palette ─────────────────────────────────────────────────────

NAVY   = RGBColor(0x1E, 0x3A, 0x5F)
BLUE   = RGBColor(0x25, 0x63, 0xEB)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
DARK   = RGBColor(0x0F, 0x17, 0x2A)
GRAY   = RGBColor(0x64, 0x74, 0x8B)
LGRAY  = RGBColor(0x94, 0xA3, 0xB8)
GREEN  = RGBColor(0x05, 0x96, 0x69)
AMBER  = RGBColor(0xD9, 0x77, 0x06)
RED    = RGBColor(0xDC, 0x26, 0x26)
SUCCESS = RGBColor(0x05, 0x96, 0x69)
WARN    = RGBColor(0xD9, 0x77, 0x06)
DANGER  = RGBColor(0xDC, 0x26, 0x26)

NAVY_HEX   = '#1E3A5F'
BLUE_HEX   = '#2563EB'
GRAY_HEX   = '#64748B'
LGRAY_HEX  = '#94A3B8'

# Number formatting
def fmt_num(n):
    return f'{int(n):,}' if n == int(n) else f'{n:,.1f}'

def fmt_money(n):
    n = float(n or 0)
    if abs(n) >= 10_000_000:
        return f'INR {n/10_000_000:.1f} Cr'
    if abs(n) >= 100_000:
        return f'INR {n/100_000:.1f} L'
    return f'INR {int(n):,}'

def fmt_date(d):
    if isinstance(d, str) and d:
        try:
            return datetime.strptime(d[:10], '%Y-%m-%d').strftime('%d %b %Y')
        except ValueError:
            return d
    return '-'

# ── Routes ─────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return send_file(str(HERE / 'index.html'))

@app.route('/api/generate', methods=['POST'])
def api_generate():
    """Generate reports from uploaded CSV files.
    Accepts multipart form with:
      - files: one or more CSV files
      - format: 'ppt' | 'word' | 'pdf'
    Returns a ZIP file containing the generated report(s).
    """
    files = request.files.getlist('files')
    fmt = request.form.get('format', 'ppt')

    if not files:
        return jsonify({'error': 'No CSV files uploaded'}), 400

    csv_texts = {}
    for f in files:
        csv_texts[f.filename] = f.read().decode('utf-8', errors='replace')

    datasets = []
    for fname, text in csv_texts.items():
        records = parse_csv(text)
        if records:
            datasets.append({
                'filename': fname,
                'records': records,
                'client': extract_client_name(records, fname),
            })

    if not datasets:
        return jsonify({'error': 'No valid records found in CSVs'}), 400

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for ds in datasets:
            report = build_report(ds['records'])
            if fmt == 'ppt':
                ppt_bytes = generate_ppt(report)
                zf.writestr(f"{safe_filename(ds['client'])}_Logistics_Report.pptx", ppt_bytes)
            elif fmt == 'word':
                docx_bytes = generate_docx(report)
                zf.writestr(f"{safe_filename(ds['client'])}_Logistics_Report.docx", docx_bytes)
            else:
                pdf_bytes = generate_pdf(report)
                zf.writestr(f"{safe_filename(ds['client'])}_Logistics_Report.pdf", pdf_bytes)

        if len(datasets) >= 2:
            all_recs = []
            for ds in datasets:
                all_recs.extend(ds['records'])
            merged_report = build_report(all_recs)
            merged_report['client'] = 'All Customers (Consolidated)'
            if fmt == 'ppt':
                zf.writestr('All_Customers_Consolidated_Logistics_Report.pptx', generate_ppt(merged_report))
            elif fmt == 'word':
                zf.writestr('All_Customers_Consolidated_Logistics_Report.docx', generate_docx(merged_report))
            else:
                zf.writestr('All_Customers_Consolidated_Logistics_Report.pdf', generate_pdf(merged_report))

    buf.seek(0)
    return send_file(
        buf,
        mimetype='application/zip',
        as_attachment=True,
        download_name=f'Logistics_Reports_{date.today().isoformat()}.zip'
    )

# ── CSV parsing ────────────────────────────────────────────────────────

def parse_csv(text):
    text = text.lstrip('\ufeff')
    lines = text.split('\n')
    if not lines:
        return []
    header = [c.strip().strip('"').strip("'") for c in lines[0].split(',')]
    records = []
    for line in lines[1:]:
        line = line.strip()
        if not line:
            continue
        vals = parse_csv_line(line)
        if len(vals) != len(header):
            vals = (vals + [''] * len(header))[:len(header)]
        rec = dict(zip(header, vals))
        records.append(rec)
    return records

def parse_csv_line(line):
    result = []
    current = ''
    quoted = False
    for ch in line:
        if quoted:
            if ch == '"':
                quoted = False
            else:
                current += ch
        elif ch == '"':
            quoted = True
        elif ch == ',':
            result.append(current.strip())
            current = ''
        else:
            current += ch
    result.append(current.strip())
    return result

def extract_client_name(records, filename):
    sample = records[0] if records else {}
    client_key = None
    for k in sample:
        kl = k.lower().strip()
        if kl in ('client', 'customer code', 'customer'):
            client_key = k
            break
    if client_key:
        vals = {r.get(client_key, '') for r in records if r.get(client_key, '')}
        vals = {v for v in vals if v.strip()}
        if len(vals) == 1:
            return list(vals)[0].replace('`', '').strip().title()
    base = filename.replace('.csv', '').replace('_', ' ').replace('-', ' ').title()
    return base

def clean_val(v):
    return str(v or '').replace('`', '').strip()

def parse_num(v):
    try:
        return float(clean_val(v).replace(',', ''))
    except ValueError:
        return 0.0

def parse_dt(v):
    s = clean_val(v)
    if not s:
        return None
    m = re.match(r'(\d{4})-(\d{2})-(\d{2})', s)
    if m:
        return datetime(int(m[1]), int(m[2]), int(m[3]))
    return None

def guess_state(city):
    city_lower = (city or '').lower()
    mapping = {
        'bengaluru': 'Karnataka', 'bangalore': 'Karnataka',
        'mumbai': 'Maharashtra', 'pune': 'Maharashtra', 'nagpur': 'Maharashtra', 'thane': 'Maharashtra',
        'delhi': 'Delhi', 'new delhi': 'Delhi',
        'hyderabad': 'Telangana',
        'chennai': 'Tamil Nadu',
        'kolkata': 'West Bengal',
        'ahmedabad': 'Gujarat', 'jaipur': 'Rajasthan',
        'lucknow': 'Uttar Pradesh', 'chandigarh': 'Chandigarh',
    }
    for k, v in mapping.items():
        if k in city_lower:
            return v
    return city or 'Unknown'

# ── Report data builder ─────────────────────────────────────────────────

def build_report(records):
    normalized = []
    for r in records:
        raw = {k.lower().strip(): v for k, v in r.items()}
        n = {}
        n['id'] = clean_val(raw.get('order id', raw.get('reference number', raw.get('lrn', ''))))
        n['client'] = clean_val(raw.get('client', raw.get('customer code', raw.get('customer', ''))))
        n['boxes'] = parse_num(raw.get('no of boxes', raw.get('num pieces', '0')))
        n['origin'] = clean_val(raw.get('origin city', raw.get('sender city', ''))).title()
        n['dest'] = clean_val(raw.get('destination city', raw.get('consignee city', ''))).title()
        warehouse = clean_val(raw.get('client location/warehouse', raw.get('origin hub name', '')))
        n['warehouse'] = warehouse.title() if warehouse else 'Unassigned'
        n['manifest'] = parse_dt(raw.get('manifest date', raw.get('created at', '')))
        n['pickup'] = parse_dt(raw.get('pickup date', raw.get('last pickup completed time', '')))
        n['promise'] = parse_dt(raw.get('promise date', raw.get('expected delivery date', raw.get('expected date', ''))))
        n['delivered'] = parse_dt(raw.get('delivered date', raw.get('delivered time', '')))
        status = clean_val(raw.get('current status', raw.get('status', '')))
        n['status'] = status
        n['is_delivered'] = status.lower() in ('delivered', 'delivered to consignee') or n['delivered'] is not None
        n['state'] = clean_val(raw.get('state', '')).title()
        if not n['state'] or n['state'] == 'Unknown':
            n['state'] = guess_state(n['dest'])
        n['amount'] = parse_num(raw.get('package amount', raw.get('declared value', '0')))
        n['weight'] = parse_num(raw.get('weight', '0'))
        n['attempts'] = int(parse_num(raw.get('attempt count', '0')))
        n['last_scan'] = parse_dt(raw.get('last scan date', ''))
        n['remarks'] = clean_val(raw.get('remarks', raw.get('delivery failure reason', '')))
        n['zone'] = clean_val(raw.get('invoice zone', ''))
        normalized.append(n)

    shipments = len(normalized)
    delivered = sum(1 for r in normalized if r['is_delivered'])
    open_s = shipments - delivered
    late_delivered = sum(1 for r in normalized if r['is_delivered'] and r['promise'] and r['delivered'] and r['delivered'] > r['promise'])
    open_delayed = sum(1 for r in normalized if not r['is_delivered'] and r['promise'] and r['promise'] < datetime.now())

    delivered_with_promise = [r for r in normalized if r['is_delivered'] and r['delivered'] and r['promise']]
    on_time = sum(1 for r in delivered_with_promise if r['delivered'] <= r['promise'])
    on_time_rate = round(on_time / len(delivered_with_promise) * 100) if delivered_with_promise else 0

    tat_records = [r for r in normalized if r['is_delivered'] and r['delivered'] and (r['pickup'] or r['manifest'])]
    avg_tat = 0
    if tat_records:
        days = []
        for r in tat_records:
            start = r['pickup'] or r['manifest']
            days.append((r['delivered'] - start).days)
        avg_tat = sum(days) / len(days) if days else 0

    total_boxes = sum(r['boxes'] for r in normalized)
    total_weight = sum(r['weight'] for r in normalized)
    total_amount = sum(r['amount'] for r in normalized)
    attempted = sum(1 for r in normalized if r['attempts'] > 0)
    delivered_rate = round(delivered / shipments * 100) if shipments else 0

    # status breakdown
    status_counts = {}
    for r in normalized:
        s = r['status'] or 'Unknown'
        status_counts[s] = status_counts.get(s, 0) + 1
    status_entries = sorted(status_counts.items(), key=lambda x: -x[1])[:6]

    # state counts
    state_counts = {}
    for r in normalized:
        st = r['state'] or 'Unknown'
        state_counts[st] = state_counts.get(st, 0) + 1
    state_entries = sorted(state_counts.items(), key=lambda x: -x[1])[:8]

    # lanes
    lane_counts = {}
    for r in normalized:
        lane = f"{r['origin']} -> {r['dest']}"
        lane_counts[lane] = lane_counts.get(lane, 0) + 1
    lane_entries = sorted(lane_counts.items(), key=lambda x: -x[1])[:10]

    # aging
    aging = {'0-2d': 0, '3-5d': 0, '6-8d': 0, '9d+': 0}
    now = datetime.now()
    for r in normalized:
        if not r['is_delivered']:
            start = r['pickup'] or r['manifest'] or r['last_scan']
            if start:
                age = max(0, (now - start).days)
                if age <= 2:
                    aging['0-2d'] += 1
                elif age <= 5:
                    aging['3-5d'] += 1
                elif age <= 8:
                    aging['6-8d'] += 1
                else:
                    aging['9d+'] += 1

    # priority records
    priority = sorted(normalized, key=lambda r: (
        0 if (not r['is_delivered'] and r['promise'] and r['promise'] < now) else
        1 if not r['is_delivered'] else 2,
        r['promise'] if r['promise'] else datetime.max
    ))

    top_state = state_entries[0][0] if state_entries else 'No state'
    top_lane = lane_entries[0][0] if lane_entries else 'No lane'
    oldest_bucket = max(aging, key=aging.get)

    insights = [
        f"{delivered_rate}% delivery closure across {fmt_num(shipments)} shipments; {fmt_num(open_s)} remain open.",
        f"{on_time_rate}% on-time delivery performance with {fmt_num(late_delivered)} late delivered shipments.",
        f"{fmt_num(open_delayed)} open shipments past promise date — require ETA updates.",
        f"{top_state} is the largest destination state with {fmt_num(state_entries[0][1])} shipments." if state_entries else '',
        f"Highest volume lane: {top_lane} — {fmt_num(lane_entries[0][1])} shipments." if lane_entries else '',
    ]

    return {
        'client': clean_val(normalized[0].get('client', '')).title() if normalized else 'Customer',
        'period': f"{fmt_date(normalized[0]['manifest'])} to {fmt_date(normalized[-1]['manifest'])}" if normalized and normalized[0]['manifest'] and normalized[-1]['manifest'] else 'Selected period',
        'generated': f"Generated {datetime.now().strftime('%d %b %Y %H:%M')}",
        'shipments': shipments,
        'delivered': delivered,
        'open': open_s,
        'late_delivered': late_delivered,
        'open_delayed': open_delayed,
        'delivered_rate': delivered_rate,
        'on_time_rate': on_time_rate,
        'avg_tat': avg_tat,
        'avg_tat_label': f'{avg_tat:.1f}d',
        'total_boxes': total_boxes,
        'total_weight': total_weight,
        'total_amount': total_amount,
        'attempted': attempted,
        'value_label': fmt_money(total_amount),
        'weight_label': f'{total_weight:.1f} kg',
        'status_entries': status_entries,
        'state_entries': state_entries,
        'lane_entries': lane_entries,
        'aging': aging,
        'priority_records': priority[:14],
        'insights': [i for i in insights if i],
        'records': normalized,
    }

def safe_filename(name):
    return re.sub(r'[\\/:*?"<>|]+', '-', str(name or 'Report')).strip().replace(' ', '_')[:60]

# ══════════════════════════════════════════════════════════════════════
#  PPT Generator (python-pptx)
# ══════════════════════════════════════════════════════════════════════

def generate_ppt(report):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    # ── helpers ──────────────────────────────────────────────────────
    sw = Inches(13.33)
    sh = Inches(7.5)
    M = Inches(0.5)

    def add_rect(slide, left, top, width, height, fill, line=None):
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
        if line:
            shape.line.color.rgb = line
            shape.line.width = Pt(0.7)
        else:
            shape.line.fill.background()
        return shape

    def add_text(slide, left, top, width, height, text, size=12, bold=False, color=DARK, align=PP_ALIGN.LEFT, font='Aptos'):
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = str(text)
        p.font.size = Pt(size)
        p.font.bold = bold
        p.font.color.rgb = color
        p.font.name = font
        p.alignment = align
        return txBox

    def add_multiline(slide, left, top, width, height, lines, size=10, color=DARK, bold=False):
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.word_wrap = True
        for i, line in enumerate(lines):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            p.text = str(line)
            p.font.size = Pt(size)
            p.font.color.rgb = color
            p.font.name = 'Aptos'
            p.space_after = Pt(4)

    def add_kpi_row(slide, metrics, y):
        card_w = Inches(2.82)
        card_h = Inches(0.98)
        gap = Inches(0.3)
        colors_fill = [RGBColor(0xEB, 0xF5, 0xFF), RGBColor(0xD1, 0xFA, 0xE5), RGBColor(0xFE, 0xF3, 0xC7), RGBColor(0xFE, 0xE2, 0xE2)]
        colors_border = [RGBColor(0x93, 0xC5, 0xFD), RGBColor(0x6E, 0xE7, 0xB7), RGBColor(0xFC, 0xD3, 0x4D), RGBColor(0xFC, 0xA5, 0xA5)]
        for i, (label, value, sub) in enumerate(metrics):
            x = M + i * (card_w + gap)
            add_rect(slide, x, y, card_w, card_h, colors_fill[i % 4], colors_border[i % 4])
            add_text(slide, x + Inches(0.12), y + Inches(0.08), Inches(2.58), Inches(0.3),
                     str(value), 18, True, DARK)
            add_text(slide, x + Inches(0.12), y + Inches(0.46), Inches(2.58), Inches(0.17),
                     label, 7.5, True, RGBColor(0x47, 0x55, 0x69))
            add_text(slide, x + Inches(0.12), y + Inches(0.68), Inches(2.58), Inches(0.18),
                     sub or '', 7, False, GRAY)

    def add_table(slide, rows, x, y, w, h, col_weights):
        if not rows:
            return
        row_h = h / len(rows)
        total_w = sum(col_weights)
        for ri, row in enumerate(rows):
            cx = x
            for ci, cell_val in enumerate(row):
                cw = int(w * col_weights[ci] / total_w)
                is_header = ri == 0
                fill = NAVY if is_header else (WHITE if ri % 2 == 0 else RGBColor(0xF1, 0xF5, 0xF9))
                border = NAVY if is_header else RGBColor(0xDD, 0xE7, 0xF0)
                add_rect(slide, cx, y + ri * row_h, cw, row_h, fill, border)
                fs = 7 if is_header else 6.5
                add_text(slide, cx + Inches(0.04), y + ri * row_h + Inches(0.03),
                         cw - Inches(0.08), row_h - Inches(0.06),
                         str(cell_val or ''), fs, is_header, WHITE if is_header else RGBColor(0x1F, 0x29, 0x37))
                cx += cw

    def add_chart_placeholder(slide, x, y, w, h, label):
        add_rect(slide, x - Inches(0.02), y - Inches(0.3), w + Inches(0.04), h + Inches(0.38),
                 WHITE, RGBColor(0xDD, 0xE7, 0xF0))
        add_text(slide, x + Inches(0.06), y - Inches(0.26), w - Inches(0.12), Inches(0.2),
                 label, 9, True, NAVY)
        add_rect(slide, x, y, w, h, RGBColor(0xF8, 0xFA, 0xFC), RGBColor(0xDD, 0xE7, 0xF0))
        add_text(slide, x, y + h // 2 - Inches(0.2), w, Inches(0.4),
                 'Chart data', 10, False, GRAY, PP_ALIGN.CENTER)

    def add_page_header(slide, title, subtitle='', description=''):
        add_rect(slide, 0, 0, sw, Inches(0.08), NAVY)
        add_text(slide, M, Inches(0.18), Inches(7.8), Inches(0.34), title, 20, True, DARK)
        if subtitle:
            add_text(slide, M, Inches(0.56), Inches(8.8), Inches(0.2), subtitle, 9, False, GRAY)
        if description:
            add_text(slide, M, Inches(0.78), Inches(8.8), Inches(0.17), description, 8, False, LGRAY)

    def add_footer(slide, text, client):
        add_rect(slide, 0, Inches(6.92), sw, Inches(0.01), RGBColor(0xCB, 0xD5, 0xE1))
        add_text(slide, M, Inches(6.95), Inches(10.6), Inches(0.16), text or '', 7, False, LGRAY)
        add_text(slide, Inches(11.25), Inches(6.95), Inches(1.65), Inches(0.16),
                 client or 'Logistics', 7, True, NAVY, PP_ALIGN.RIGHT)

    def add_callout(slide, x, y, w, h, title, bullets, accent=BLUE):
        add_rect(slide, x, y, w, h, RGBColor(0xF8, 0xFA, 0xFC), RGBColor(0xDD, 0xE7, 0xF0))
        add_rect(slide, x, y, Inches(0.06), h, accent, accent)
        add_text(slide, x + Inches(0.2), y + Inches(0.12), w - Inches(0.36), Inches(0.2),
                 title, 9, True, DARK)
        lines = [f'- {b}' for b in bullets[:5]]
        add_multiline(slide, x + Inches(0.24), y + Inches(0.4), w - Inches(0.48), h - Inches(0.5),
                      lines, 8.2, RGBColor(0x33, 0x41, 0x55))

    # ── Slide 1: Cover ───────────────────────────────────────────────
    sl = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    add_rect(sl, 0, 0, sw, Inches(0.55), NAVY)
    add_text(sl, Inches(0.5), Inches(0.08), Inches(12.3), Inches(0.4),
             'MONTHLY LOGISTICS REPORT', 10, True, WHITE)
    add_rect(sl, M, Inches(1.0), Inches(12.1), Inches(0.03), BLUE)
    add_text(sl, M, Inches(1.28), Inches(8.5), Inches(0.55),
             'Logistics Performance Report', 28, True, DARK)
    add_text(sl, M + Inches(0.02), Inches(1.9), Inches(7.5), Inches(0.32),
             report['client'], 16, True, NAVY)
    add_text(sl, M + Inches(0.02), Inches(2.28), Inches(7.5), Inches(0.25),
             report['period'], 11, False, GRAY)
    add_rect(sl, M, Inches(2.75), Inches(12.1), Inches(0.02), RGBColor(0xCB, 0xD5, 0xE1))
    add_kpi_row(sl, [
        ('Total Shipments', fmt_num(report['shipments']), 'All consignments'),
        ('Delivered', f"{report['delivered_rate']}%", f"{fmt_num(report['delivered'])} closed"),
        ('On-Time Rate', f"{report['on_time_rate']}%", f"{fmt_num(report['late_delivered'])} late"),
        ('Open Delayed', fmt_num(report['open_delayed']), f"{fmt_num(report['open'])} open"),
    ], Inches(3.15))
    add_callout(sl, Inches(0.72), Inches(4.65), Inches(11.9), Inches(1.55),
                'Executive Summary', report['insights'][:3])
    add_text(sl, M, Inches(6.94), Inches(10.8), Inches(0.18),
             report['generated'], 7.5, False, LGRAY)
    add_footer(sl, 'Confidential — For internal use only', report['client'])

    # ── Slide 2: Executive Summary ───────────────────────────────────
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    add_page_header(sl, 'Executive Summary', report['period'], 'Key metrics, trends, and aging overview')
    add_kpi_row(sl, [
        ('Package Value', report['value_label'], 'Declared invoice value'),
        ('Total Boxes', fmt_num(report['total_boxes']), report['weight_label']),
        ('Avg TAT', report['avg_tat_label'], 'Pickup to delivery'),
        ('Attempted', fmt_num(report['attempted']), 'Shipments with attempt'),
    ], Inches(0.95))

    left_x = M + Inches(0.05)
    right_x = Inches(6.75)
    tw = Inches(5.75)
    add_callout(sl, left_x, Inches(2.22), tw, Inches(2.15), 'Key Insights', report['insights'])
    status_rows = [['Status', 'Count', 'Share', 'Comment']]
    total_s = max(1, report['shipments'])
    for s, c in report['status_entries']:
        share = round(c / total_s * 100)
        comment = 'Closed' if s.lower() == 'delivered' else ('Moving' if 'transit' in s.lower() else 'Monitor')
        status_rows.append([s, fmt_num(c), f'{share}%', comment])
    add_table(sl, status_rows, right_x, Inches(2.22), tw, Inches(2.15), [2.4, 1.05, 1.05, 1.25])

    add_chart_placeholder(sl, left_x, Inches(4.82), tw, Inches(1.75), 'Daily Manifest Trend')
    aging_rows = [['Open Aging', 'Count', 'Interpretation']]
    for label, val in sorted(report['aging'].items()):
        note = 'Escalate' if label == '9d+' else ('Watchlist' if label == '6-8d' else 'Follow-up')
        aging_rows.append([label, fmt_num(val), note])
    add_table(sl, aging_rows, right_x, Inches(4.82), tw, Inches(1.75), [2.2, 1.2, 2.35])
    add_footer(sl, report['generated'], report['client'])

    # ── Slide 3: Performance ─────────────────────────────────────────
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    add_page_header(sl, 'Service Performance', 'Delivery closure, promise adherence, and daily movement',
                    'Operational metrics with visual breakdowns')
    add_chart_placeholder(sl, Inches(0.5), Inches(1.02), Inches(5.95), Inches(4.35), 'Shipment Status Distribution')
    add_chart_placeholder(sl, Inches(6.8), Inches(1.02), Inches(5.95), Inches(2.35), 'Daily Manifest Volume')
    add_callout(sl, Inches(6.8), Inches(3.78), Inches(5.95), Inches(1.6), 'Performance Notes', [
        f"{report['delivered_rate']}% of shipments closed delivered.",
        f"{report['on_time_rate']}% met promise date.",
        f"{fmt_num(report['open_delayed'])} open shipments past promise — review first.",
    ], WARN)
    add_footer(sl, report['generated'], report['client'])

    # ── Slide 4: Network ─────────────────────────────────────────────
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    add_page_header(sl, 'Network & Movement', 'Geographic distribution, aging profile, and top lanes',
                    'Destination states, open aging, and lane volumes')
    add_chart_placeholder(sl, Inches(0.5), Inches(1.02), Inches(5.95), Inches(3.25), 'Top Destination States')
    add_chart_placeholder(sl, Inches(6.8), Inches(1.02), Inches(5.95), Inches(3.25), 'Open Shipment Aging')
    lane_rows = [['Lane', 'Shipments', 'Delivered', 'Open', 'Rate']]
    for lane_name, count in report['lane_entries'][:6]:
        lane_recs = [r for r in report['records'] if f"{r['origin']} -> {r['dest']}" == lane_name]
        lane_del = sum(1 for r in lane_recs if r['is_delivered'])
        rate = round(lane_del / count * 100) if count else 0
        lane_rows.append([lane_name[:46], fmt_num(count), fmt_num(lane_del), fmt_num(count - lane_del), f'{rate}%'])
    add_table(sl, lane_rows, Inches(0.55), Inches(4.82), Inches(12.1), Inches(1.65), [5.3, 1.65, 1.65, 1.65, 1.85])
    add_footer(sl, report['generated'], report['client'])

    # ── Slide 5: Exceptions ──────────────────────────────────────────
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    add_page_header(sl, 'Exceptions & Priority Items', 'Delayed, late, and open shipments',
                    'Actionable items for customer review')
    add_callout(sl, Inches(0.55), Inches(0.98), Inches(12.1), Inches(0.88),
                'Recommended Actions', [
                    'Prioritize open delayed shipments — confirm ETAs with customers immediately.',
                    'Review late-delivered lanes to identify recurring network or appointment issues.',
                ], DANGER)
    exc_rows = [['Order', 'Destination', 'Status', 'Promise', 'Last Scan', 'Amount', 'Action']]
    now = datetime.now()
    for r in report['priority_records'][:8]:
        is_delayed = not r['is_delivered'] and r['promise'] and r['promise'] < now
        is_late = r['is_delivered'] and r['promise'] and r['delivered'] and r['delivered'] > r['promise']
        if not r['is_delivered'] or is_delayed or is_late:
            status_label = 'Delayed' if is_delayed else ('Delivered' if r['is_delivered'] else 'Open')
            action = 'Expedite / confirm ETA' if is_delayed else ('Review late reason' if r['is_delivered'] else 'Track movement')
            exc_rows.append([
                r['id'][:22], f"{r['dest']}, {r['state']}"[:28], status_label,
                fmt_date(r['promise']), fmt_date(r['last_scan']), fmt_money(r['amount']), action,
            ])
    add_table(sl, exc_rows, Inches(0.55), Inches(2.18), Inches(12.1), Inches(4.2),
              [2.05, 2.4, 1.35, 1.35, 1.65, 1.45, 1.85])
    add_footer(sl, report['generated'], report['client'])

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════
#  DOCX Generator (python-docx)
# ══════════════════════════════════════════════════════════════════════

def generate_docx(report):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Aptos'
    style.font.size = DocPt(10)
    style.paragraph_format.space_after = DocPt(4)

    # ── Title ────────────────────────────────────────────────────────
    p = doc.add_heading('LOGISTICS PERFORMANCE REPORT', level=0)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in p.runs:
        run.font.color.rgb = DocRGB(0x1E, 0x3A, 0x5F)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(report['client'])
    run.bold = True
    run.font.size = DocPt(16)
    run.font.color.rgb = DocRGB(0x25, 0x63, 0xEB)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(report['period'])
    run.font.size = DocPt(10)
    run.font.color.rgb = DocRGB(0x64, 0x74, 0x8B)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(report['generated'])
    run.font.size = DocPt(8)
    run.font.color.rgb = DocRGB(0x94, 0xA3, 0xB8)

    doc.add_paragraph().add_run().add_break()

    # ── Executive Summary ────────────────────────────────────────────
    doc.add_heading('Executive Summary', level=1)
    p = doc.add_paragraph()
    run = p.add_run(
        f"Total Shipments: {fmt_num(report['shipments'])}  |  "
        f"Delivered: {report['delivered_rate']}%  |  "
        f"On-Time: {report['on_time_rate']}%  |  "
        f"Open Delayed: {fmt_num(report['open_delayed'])}"
    )
    run.font.size = DocPt(10)
    run.font.color.rgb = DocRGB(0x33, 0x33, 0x33)

    for insight in report['insights']:
        p = doc.add_paragraph(style='List Bullet')
        run = p.add_run(insight)
        run.font.size = DocPt(9)
        run.font.color.rgb = DocRGB(0x47, 0x55, 0x69)

    # ── Tables ───────────────────────────────────────────────────────
    sections_data = [
        ('Status Breakdown', report['status_entries'], ['Status', 'Count', 'Share', 'Comment'],
         lambda s, c: [s, fmt_num(c), f"{round(c / max(1, report['shipments']) * 100)}%",
                       'Closed' if s.lower() == 'delivered' else ('Moving' if 'transit' in s.lower() else 'Monitor')]),
        ('Open Aging', sorted(report['aging'].items()), ['Aging', 'Count', 'Interpretation'],
         lambda label, val: [label, fmt_num(val),
                             'Escalate' if label == '9d+' else ('Watchlist' if label == '6-8d' else 'Follow-up')]),
        ('Top Lanes', report['lane_entries'][:6], ['Lane', 'Shipments', 'Delivered', 'Open', 'Rate'],
         lambda lane, count: [lane[:50], fmt_num(count),
                              fmt_num(sum(1 for r in report['records'] if r['is_delivered'] and f"{r['origin']} -> {r['dest']}" == lane)),
                              '—', '—']),
    ]

    for title, entries, headers, row_fn in sections_data:
        doc.add_heading(title, level=2)
        table = doc.add_table(rows=1 + len(entries), cols=len(headers))
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        for i, h in enumerate(headers):
            cell = table.rows[0].cells[i]
            cell.text = h
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.bold = True
                    run.font.size = DocPt(9)
                    run.font.color.rgb = DocRGB(0x1E, 0x3A, 0x5F)
            from docx.oxml.ns import qn
            shading = cell._element.get_or_add_tcPr()
            shading_elem = shading.makeelement(qn('w:shd'), {
                qn('w:fill'): 'EBF5FF',
                qn('w:val'): 'clear'
            })
            shading.append(shading_elem)

        for ri, (key, val) in enumerate(entries):
            cells = table.rows[ri + 1].cells
            row_data = row_fn(key, val)
            for ci, text in enumerate(row_data):
                cells[ci].text = str(text)
                for paragraph in cells[ci].paragraphs:
                    for run in paragraph.runs:
                        run.font.size = DocPt(8)

        doc.add_paragraph()

    # ── Exceptions Table ─────────────────────────────────────────────
    doc.add_heading('Priority Exceptions', level=2)
    now = datetime.now()
    exc_headers = ['Order', 'Destination', 'Status', 'Promise', 'Last Scan', 'Amount', 'Action']
    exc_data = []
    for r in report['priority_records'][:8]:
        is_delayed = not r['is_delivered'] and r['promise'] and r['promise'] < now
        if not r['is_delivered'] or is_delayed:
            status_label = 'Delayed' if is_delayed else ('Delivered' if r['is_delivered'] else 'Open')
            action = 'Expedite / confirm ETA' if is_delayed else ('Review late reason' if r['is_delivered'] else 'Track movement')
            exc_data.append([
                r['id'][:22], f"{r['dest']}, {r['state']}"[:28], status_label,
                fmt_date(r['promise']), fmt_date(r['last_scan']), fmt_money(r['amount']), action,
            ])

    if exc_data:
        table = doc.add_table(rows=1 + len(exc_data), cols=len(exc_headers))
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, h in enumerate(exc_headers):
            cell = table.rows[0].cells[i]
            cell.text = h
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.bold = True
                    run.font.size = DocPt(9)
                    run.font.color.rgb = DocRGB(0x1E, 0x3A, 0x5F)
            from docx.oxml.ns import qn
            shading = cell._element.get_or_add_tcPr()
            shading_elem = shading.makeelement(qn('w:shd'), {
                qn('w:fill'): 'FFF0F0',
                qn('w:val'): 'clear'
            })
            shading.append(shading_elem)

        for ri, row_data in enumerate(exc_data):
            for ci, text in enumerate(row_data):
                table.rows[ri + 1].cells[ci].text = str(text)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════
#  PDF Generator (fpdf2)
# ══════════════════════════════════════════════════════════════════════

def safe_text(t):
    """Replace special chars not supported by fpdf2 latin-1 fonts."""
    replacements = {
        '\u2014': '--', '\u2013': '-', '\u2018': "'", '\u2019': "'",
        '\u201c': '"', '\u201d': '"', '\u2026': '...', '\u2022': '-',
        '\u00b0': ' deg', '\u2122': '(TM)', '\u00ae': '(R)',
        '\u00a0': ' ', '\u20b9': 'Rs.', '\ufffd': '',
    }
    t = str(t)
    for old, new in replacements.items():
        t = t.replace(old, new)
    return t.encode('latin-1', errors='replace').decode('latin-1')

class LogisticsPDF(FPDF):
    def header(self):
        self.set_font('Helvetica', 'B', 8)
        self.set_text_color(0x1E, 0x3A, 0x5F)
        self.cell(0, 6, 'LOGISTICS PERFORMANCE REPORT', align='C')
        self.ln(8)
        self.set_draw_color(0x25, 0x63, 0xEB)
        self.set_line_width(0.5)
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(4)

    def footer(self):
        self.set_y(-15)
        self.set_font('Helvetica', '', 7)
        self.set_text_color(0x94, 0xA3, 0xB8)
        self.cell(0, 10, f'Page {self.page_no()}/{{nb}}', align='C')

    def section_title(self, title):
        self.set_font('Helvetica', 'B', 13)
        self.set_text_color(0x1E, 0x3A, 0x5F)
        self.cell(0, 8, title)
        self.ln(6)
        self.set_draw_color(0x25, 0x63, 0xEB)
        self.set_line_width(0.3)
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(4)

    def sub_title(self, title):
        self.set_font('Helvetica', 'B', 10)
        self.set_text_color(0x1E, 0x3A, 0x5F)
        self.cell(0, 7, title)
        self.ln(5)

    def body_text(self, text, size=9, color=(0x33, 0x33, 0x33)):
        self.set_font('Helvetica', '', size)
        self.set_text_color(*color)
        self.multi_cell(0, 5, safe_text(text))
        self.ln(1)

    def kpi_card(self, x, y, w, h, value, label, sub, bg, border):
        self.set_fill_color(*bg)
        self.set_draw_color(*border)
        self.set_line_width(0.3)
        self.rect(x, y, w, h, 'DF')
        self.set_xy(x + 2, y + 1.5)
        self.set_font('Helvetica', 'B', 14)
        self.set_text_color(0x0F, 0x17, 0x2A)
        self.cell(w - 4, 6, safe_text(str(value)))
        self.set_xy(x + 2, y + 8)
        self.set_font('Helvetica', 'B', 6.5)
        self.set_text_color(0x47, 0x55, 0x69)
        self.cell(w - 4, 4, safe_text(label))
        self.set_xy(x + 2, y + 12.5)
        self.set_font('Helvetica', '', 6)
        self.set_text_color(0x64, 0x74, 0x8B)
        self.cell(w - 4, 4, safe_text(sub))

    def data_table(self, headers, rows, col_widths=None, header_bg=(0x1E, 0x3A, 0x5F)):
        if not rows:
            return
        if col_widths is None:
            col_widths = [190 // len(headers)] * len(headers)
        total_w = sum(col_widths)
        scale = 190 / total_w
        col_widths = [w * scale for w in col_widths]

        # header
        self.set_fill_color(*header_bg)
        self.set_text_color(255, 255, 255)
        self.set_font('Helvetica', 'B', 6.5)
        for i, h in enumerate(headers):
            self.cell(col_widths[i], 6, h, border=1, fill=True, align='C')
        self.ln()

        # rows
        for ri, row in enumerate(rows):
            if ri % 2 == 0:
                self.set_fill_color(255, 255, 255)
            else:
                self.set_fill_color(0xF8, 0xFA, 0xFC)
            self.set_text_color(0x33, 0x33, 0x33)
            self.set_font('Helvetica', '', 6.5)
            max_h = 5
            for ci, cell_val in enumerate(row):
                text = safe_text(str(cell_val or ''))[:60]
                self.cell(col_widths[ci], max_h, text, border=1, fill=True)
            self.ln()


def generate_pdf(report):
    pdf = LogisticsPDF()
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.add_page()

    # ── Title block ──────────────────────────────────────────────────
    pdf.set_font('Helvetica', 'B', 20)
    pdf.set_text_color(0x1E, 0x3A, 0x5F)
    pdf.cell(0, 10, 'LOGISTICS PERFORMANCE REPORT', align='C')
    pdf.ln(12)

    pdf.set_font('Helvetica', 'B', 12)
    pdf.set_text_color(0x25, 0x63, 0xEB)
    pdf.cell(0, 7, report['client'], align='C')
    pdf.ln(7)

    pdf.set_font('Helvetica', '', 8)
    pdf.set_text_color(0x64, 0x74, 0x8B)
    pdf.cell(0, 5, report['period'], align='C')
    pdf.ln(5)

    pdf.set_font('Helvetica', '', 7)
    pdf.set_text_color(0x94, 0xA3, 0xB8)
    pdf.cell(0, 5, report['generated'], align='C')
    pdf.ln(10)

    # ── KPI row ──────────────────────────────────────────────────────
    cards = [
        (fmt_num(report['shipments']), 'Total Shipments', 'All consignments',
         (0xEB, 0xF5, 0xFF), (0x93, 0xC5, 0xFD)),
        (f"{report['delivered_rate']}%", 'Delivered', f"{fmt_num(report['delivered'])} closed",
         (0xD1, 0xFA, 0xE5), (0x6E, 0xE7, 0xB7)),
        (f"{report['on_time_rate']}%", 'On-Time', f"{fmt_num(report['late_delivered'])} late",
         (0xFE, 0xF3, 0xC7), (0xFC, 0xD3, 0x4D)),
        (fmt_num(report['open_delayed']), 'Open Delayed', f"{fmt_num(report['open'])} open",
         (0xFE, 0xE2, 0xE2), (0xFC, 0xA5, 0xA5)),
    ]
    card_w = 46
    gap = 2
    start_x = 10
    for i, (val, label, sub, bg, border) in enumerate(cards):
        x = start_x + i * (card_w + gap)
        pdf.kpi_card(x, pdf.get_y(), card_w, 16, val, label, sub, bg, border)
    pdf.ln(20)

    # ── Executive Summary ────────────────────────────────────────────
    pdf.section_title('Executive Summary')
    for insight in report['insights']:
        pdf.body_text(f'- {insight}', size=8, color=(0x47, 0x55, 0x69))

    # ── Status ───────────────────────────────────────────────────────
    pdf.sub_title('Status Breakdown')
    status_rows = []
    for s, c in report['status_entries']:
        share = round(c / max(1, report['shipments']) * 100)
        comment = 'Closed' if s.lower() == 'delivered' else ('Moving' if 'transit' in s.lower() else 'Monitor')
        status_rows.append([s, fmt_num(c), f'{share}%', comment])
    pdf.data_table(['Status', 'Count', 'Share', 'Comment'], status_rows)
    pdf.ln(4)

    # ── Aging ────────────────────────────────────────────────────────
    pdf.sub_title('Open Aging')
    aging_rows = []
    for label, val in sorted(report['aging'].items()):
        note = 'Escalate' if label == '9d+' else ('Watchlist' if label == '6-8d' else 'Follow-up')
        aging_rows.append([label, fmt_num(val), note])
    pdf.data_table(['Aging', 'Count', 'Interpretation'], aging_rows)
    pdf.ln(4)

    # ── Lanes ────────────────────────────────────────────────────────
    pdf.sub_title('Top Lanes')
    lane_rows = []
    for lane, count in report['lane_entries'][:6]:
        lane_recs = [r for r in report['records'] if f"{r['origin']} -> {r['dest']}" == lane]
        lane_del = sum(1 for r in lane_recs if r['is_delivered'])
        rate = round(lane_del / count * 100) if count else 0
        lane_rows.append([lane[:50], fmt_num(count), fmt_num(lane_del), fmt_num(count - lane_del), f'{rate}%'])
    pdf.data_table(['Lane', 'Shipments', 'Delivered', 'Open', 'Rate'], lane_rows)
    pdf.ln(4)

    # ── Exceptions ───────────────────────────────────────────────────
    pdf.sub_title('Priority Exceptions')
    now = datetime.now()
    exc_rows = []
    for r in report['priority_records'][:8]:
        is_delayed = not r['is_delivered'] and r['promise'] and r['promise'] < now
        if not r['is_delivered'] or is_delayed:
            status_label = 'Delayed' if is_delayed else ('Delivered' if r['is_delivered'] else 'Open')
            action = 'Expedite / confirm ETA' if is_delayed else ('Review late reason' if r['is_delivered'] else 'Track')
            exc_rows.append([r['id'][:22], f"{r['dest']}, {r['state']}"[:28], status_label,
                             fmt_date(r['promise']), fmt_date(r['last_scan']), fmt_money(r['amount']), action])
    if exc_rows:
        pdf.data_table(['Order', 'Destination', 'Status', 'Promise', 'Last Scan', 'Amount', 'Action'],
                       exc_rows, col_widths=[28, 28, 18, 24, 24, 22, 22], header_bg=(0x8B, 0x00, 0x00))

    buf = io.BytesIO()
    pdf.output(buf)
    return buf.getvalue()


# ── Entry point ────────────────────────────────────────────────────────

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
