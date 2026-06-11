#!/usr/bin/env python3
"""
StartHub Africa Impact Dashboard Generator
=========================================
Run this script whenever you update the Excel file to regenerate the dashboard.

Usage:
    python3 generate_dashboard.py

The script reads the Excel file in the same folder and outputs dashboard.html.
Upload dashboard.html to OneDrive to share with the team.
"""

import pandas as pd
import json
import numpy as np
import io
from pathlib import Path
from datetime import date

EXCEL_FILE = Path(__file__).parent / 'StartHub Africa Impact Dashboard (2).xlsx'
OUTPUT_FILE = Path(__file__).parent / 'index.html'

COUNTRIES = ['Uganda', 'Tanzania', 'Kenya']
MONTHS    = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
MONTH_COLS_E = {
    'Jan':1,'Feb':11,'Mar':21,'Apr':31,'May':41,'Jun':51,
    'Jul':61,'Aug':71,'Sep':81,'Oct':91,'Nov':101,'Dec':111
}

# ─── Helpers ──────────────────────────────────────────────────────────────────

def sn(val):
    try:
        if val is None or (isinstance(val, float) and (np.isnan(val) or np.isinf(val))):
            return 0
        v = float(val)
        return 0 if np.isnan(v) or np.isinf(v) else round(v, 2)
    except:
        return 0

def blank_metrics():
    return {
        'trained_m':0,'trained_f':0,'trained_total':0,
        'tot_m':0,'tot_f':0,'tot_total':0,
        'biz_supported':0,'biz_started':0,
        'jobs':0,'internships':0,
        'grants_count':0,'grants_value':0,
        'revenue_baseline':0,'revenue_endline':0,'cohort_companies':0
    }

def add_metrics(a, b):
    return {k: a.get(k,0) + b.get(k,0) for k in blank_metrics()}

def fill_from_projects(totals, projects):
    """Fill zero fields in totals by summing project rows."""
    psum = blank_metrics()
    for p in projects:
        for k in blank_metrics():
            psum[k] += p['metrics'].get(k, 0)
    for k in ['tot_m','tot_f','tot_total','grants_count','grants_value',
               'internships','biz_started','jobs']:
        if totals.get(k, 0) == 0 and psum.get(k, 0) > 0:
            totals[k] = psum[k]
    return totals

def open_excel(path):
    """Open Excel file, trimming trailing null bytes (Excel save artefact)."""
    raw = Path(path).read_bytes()
    pos = raw.rfind(b'PK\x05\x06')
    if pos >= 0 and pos + 22 < len(raw):
        raw = raw[:pos + 22]
    return pd.ExcelFile(io.BytesIO(raw), engine='openpyxl')

# ─── Type E parser (2025, 2026) ───────────────────────────────────────────────

def extract_type_e_annual(row):
    def g(off):
        idx = 121 + off
        return sn(row.iloc[idx]) if idx < len(row) else 0
    return {
        'trained_m':g(0),'trained_f':g(1),'trained_total':g(2),
        'tot_m':g(3),'tot_f':g(4),'tot_total':g(5),
        'biz_supported':g(6),'biz_started':g(7),
        'jobs':g(8),'internships':g(9),
        'grants_count':g(10),'grants_value':g(11),
        'revenue_baseline':g(12),'revenue_endline':g(13),'cohort_companies':g(14)
    }

def extract_type_e_monthly(row, sc):
    def g(off):
        idx = sc + off
        return sn(row.iloc[idx]) if idx < len(row) else 0
    return {
        'trained_m':g(0),'trained_f':g(1),'trained_total':g(2),
        'tot_total':g(5),
        'biz_supported':g(6),'biz_started':g(7),
        'jobs':g(8),'internships':g(9)
    }

def parse_type_e(df):
    totals   = blank_metrics()
    countries = {c:{'totals':blank_metrics(),'monthly':{m:{} for m in MONTHS},'projects':[]} for c in COUNTRIES}
    monthly_all = {m:{} for m in MONTHS}
    current_country = None
    skip = {'Project Name','M','F','Total','nan','','NaN',
            'Young people/Entrepreneurs Trained','Training of Trainers','Businesses Supported','Businesses Started'}

    for _, row in df.iterrows():
        if len(row) < 130: continue
        col0 = '' if pd.isna(row.iloc[0]) else str(row.iloc[0]).strip()

        if col0 in COUNTRIES:
            current_country = col0; continue
        if col0 == 'Total Across Countries':
            totals = extract_type_e_annual(row)
            for m,mc in MONTH_COLS_E.items():
                monthly_all[m] = extract_type_e_monthly(row, mc)
            continue
        if col0.startswith('Total '):
            c = col0.replace('Total ','').strip()
            if c in COUNTRIES:
                countries[c]['totals'] = extract_type_e_annual(row)
                for m,mc in MONTH_COLS_E.items():
                    countries[c]['monthly'][m] = extract_type_e_monthly(row, mc)
            continue
        if col0 in skip or col0.startswith('Nr ') or col0.startswith('Value '):
            continue
        if current_country and col0:
            m = extract_type_e_annual(row)
            if any(m[k]>0 for k in ['trained_total','biz_supported','grants_count','biz_started','jobs','tot_total']):
                countries[current_country]['projects'].append({'name':col0,'metrics':m})

    all_projects = []
    for c in COUNTRIES:
        countries[c]['projects'] = [p for p in countries[c]['projects']
            if p['metrics'].get('trained_total',0)>0 or p['metrics'].get('biz_supported',0)>0
            or p['metrics'].get('grants_value',0)>0 or p['metrics'].get('tot_total',0)>0]
        all_projects.extend(countries[c]['projects'])
        countries[c]['totals'] = fill_from_projects(countries[c]['totals'], countries[c]['projects'])
    totals = fill_from_projects(totals, all_projects)
    return {'totals':totals,'countries':countries,'monthly':monthly_all}

# ─── Type D parser (2024) ─────────────────────────────────────────────────────

def extract_type_d_row(row):
    def g(i): return sn(row.iloc[i]) if len(row)>i else 0
    return {
        'trained_m':g(1),'trained_f':g(2),'trained_total':g(3),
        'tot_m':g(4),'tot_f':g(5),'tot_total':g(6),
        'biz_supported':g(7),'biz_started':g(10),
        'jobs':g(11),'internships':g(12),
        'grants_count':g(8),'grants_value':g(9),
        'revenue_baseline':g(13),'revenue_endline':g(14),'cohort_companies':g(15)
    }

def parse_type_d(df):
    totals   = blank_metrics()
    countries = {c:{'totals':blank_metrics(),'monthly':{},'projects':[]} for c in COUNTRIES}
    current_country = None

    for _, row in df.iterrows():
        col0 = '' if pd.isna(row.iloc[0]) else str(row.iloc[0]).strip()
        if col0 in ('Uganda ','Uganda'):   current_country='Uganda';  continue
        if col0 == 'Tanzania':             current_country='Tanzania'; continue
        if col0 == 'Kenya' or (len(row)>1 and not pd.isna(row.iloc[1]) and str(row.iloc[1]).strip()=='Kenya'):
            current_country='Kenya'; continue
        if col0 == 'Total Across Countries':
            totals = extract_type_d_row(row); continue
        if col0.startswith('Total '):
            c = col0.replace('Total ','').strip()
            if c in COUNTRIES: countries[c]['totals'] = extract_type_d_row(row)
            continue
        if col0 in {'','Project/Program','nan','2024','M','F','Total'}: continue
        if current_country and col0:
            m = extract_type_d_row(row)
            if any(m[k]>0 for k in ['trained_total','biz_supported','grants_value','tot_total']):
                countries[current_country]['projects'].append({'name':col0,'metrics':m})

    all_projects = []
    for c in COUNTRIES:
        all_projects.extend(countries[c]['projects'])
        countries[c]['totals'] = fill_from_projects(countries[c]['totals'], countries[c]['projects'])
    totals = fill_from_projects(totals, all_projects)
    return {'totals':totals,'countries':countries,'monthly':{}}

# ─── Type C parser (2023) ─────────────────────────────────────────────────────

def extract_type_c_row(row):
    def g(i): return sn(row.iloc[i]) if len(row)>i else 0
    return {
        'trained_m':g(1),'trained_f':g(2),'trained_total':g(3),
        'tot_m':g(4),'tot_f':g(5),'tot_total':g(6),
        'biz_supported':g(7),'biz_started':g(8),
        'jobs':0,'internships':g(11),
        'grants_count':g(9),'grants_value':g(10),
        'revenue_baseline':g(12),'revenue_endline':g(13),'cohort_companies':g(14)
    }

def parse_type_c(df):
    totals   = blank_metrics()
    countries = {c:{'totals':blank_metrics(),'monthly':{},'projects':[]} for c in COUNTRIES}
    current_country = None

    for _, row in df.iterrows():
        col0 = '' if pd.isna(row.iloc[0]) else str(row.iloc[0]).strip()
        if col0 in ('Uganda ','Uganda'):   current_country='Uganda';  continue
        if col0 == 'Tanzania':             current_country='Tanzania'; continue
        if col0 == 'Kenya':                current_country='Kenya';   continue
        if col0 == 'Total Across Countries':
            totals = extract_type_c_row(row); continue
        if col0.startswith('Total '):
            c = col0.replace('Total ','').strip()
            if c in COUNTRIES: countries[c]['totals'] = extract_type_c_row(row)
            continue
        if col0 in {'','Project/Program','nan','2023','M','F','Total'}: continue
        if col0.startswith('Young') or col0.startswith('Jan'): continue
        if current_country and col0:
            m = extract_type_c_row(row)
            if any(m[k]>0 for k in ['trained_total','biz_supported','grants_value','tot_total']):
                countries[current_country]['projects'].append({'name':col0,'metrics':m})

    all_projects = [p for cps in countries.values() for p in cps['projects']]
    totals = fill_from_projects(totals, all_projects)
    for c in COUNTRIES:
        countries[c]['totals'] = fill_from_projects(countries[c]['totals'], countries[c]['projects'])
    return {'totals':totals,'countries':countries,'monthly':{}}

# ─── Type B parser (2021, 2022) ───────────────────────────────────────────────

def extract_type_b_row(row):
    def g(i): return sn(row.iloc[i]) if len(row)>i else 0
    return {
        'trained_m':g(1),'trained_f':g(2),'trained_total':g(3),
        'tot_m':g(4),'tot_f':g(5),'tot_total':g(6),
        'biz_supported':g(7),'biz_started':0,
        'jobs':0,'internships':g(10),
        'grants_count':g(8),'grants_value':g(9),
        'revenue_baseline':0,'revenue_endline':0,'cohort_companies':0
    }

def parse_type_b(df, year_label):
    countries     = {c:{'totals':blank_metrics(),'monthly':{},'projects':[]} for c in COUNTRIES}
    current_country = 'Uganda'
    all_projects  = []

    for _, row in df.iterrows():
        col0 = '' if pd.isna(row.iloc[0]) else str(row.iloc[0]).strip()
        for c in COUNTRIES:
            if col0 == c: current_country = c; break
        if col0 in {'','Project/Program','nan',year_label,'M','F','Total','2021','2022'}: continue
        if col0.startswith('Young') or col0.startswith('Jan'): continue
        if col0 == 'Total':
            countries[current_country]['totals'] = extract_type_b_row(row); continue
        if col0.startswith('Total '):
            c = col0.replace('Total ','').strip()
            if c in COUNTRIES: countries[c]['totals'] = extract_type_b_row(row)
            continue
        if current_country and col0:
            m = extract_type_b_row(row)
            if any(m[k]>0 for k in ['trained_total','biz_supported','grants_value','tot_total']):
                all_projects.append({'name':col0,'country':current_country,'metrics':m})

    for c in COUNTRIES:
        ct = countries[c]['totals']
        c_projs = [p for p in all_projects if p['country']==c]
        if ct['trained_total'] == 0 and c_projs:
            summed = blank_metrics()
            for p in c_projs: summed = add_metrics(summed, p['metrics'])
            countries[c]['totals'] = summed
        else:
            countries[c]['totals'] = fill_from_projects(ct, c_projs)
        countries[c]['projects'] = c_projs

    grand = blank_metrics()
    for c in COUNTRIES: grand = add_metrics(grand, countries[c]['totals'])
    return {'totals':grand,'countries':countries,'monthly':{}}

# ─── Type A parser (2017-2019) ────────────────────────────────────────────────

def parse_type_a(df):
    totals   = blank_metrics()
    countries = {c:{'totals':blank_metrics(),'monthly':{},'projects':[]} for c in COUNTRIES}
    for _, row in df.iterrows():
        col0 = '' if pd.isna(row.iloc[0]) else str(row.iloc[0]).strip()
        if col0.startswith('Total Uganda') or col0 == 'Total':
            totals['trained_total'] = sn(row.iloc[3]) if len(row)>3 else 0
            totals['biz_supported'] = sn(row.iloc[7]) if len(row)>7 else 0
            countries['Uganda']['totals'] = dict(totals)
    return {'totals':totals,'countries':countries,'monthly':{}}

# ─── Aggregation ──────────────────────────────────────────────────────────────

def compute_overall_from_years(years_data):
    total = blank_metrics()
    country_totals = {c: blank_metrics() for c in COUNTRIES}
    for yr_data in years_data.values():
        t = yr_data.get('totals') or blank_metrics()
        total = add_metrics(total, t)
        for c in COUNTRIES:
            ct = (yr_data.get('countries') or {}).get(c,{}).get('totals') or blank_metrics()
            country_totals[c] = add_metrics(country_totals[c], ct)
    return {'all': total, 'countries': country_totals}

# ─── Data Extraction ──────────────────────────────────────────────────────────

def extract_all_data():
    print(f"Reading: {EXCEL_FILE}")
    xl = open_excel(EXCEL_FILE)
    sheets = xl.sheet_names
    data = {'updated': str(date.today()), 'years': {}}

    def add(label, fn, *args):
        data['years'][label] = fn(*args)
        print(f"  ✓ {label} parsed")

    if '2017-2019' in sheets:
        add('2017-19', parse_type_a, pd.read_excel(xl,'2017-2019',header=None))
    if '2020' in sheets:
        df = pd.read_excel(xl,'2020',header=None)
        ncols = df.dropna(axis=1,how='all').shape[1]
        fn = parse_type_e if ncols > 25 else parse_type_b
        add('2020', fn, df) if fn==parse_type_e else add('2020', fn, df, '2020')
        print(f"    ({ncols} cols)")
    for yr in ['2021','2022']:
        if yr in sheets:
            add(yr, parse_type_b, pd.read_excel(xl,yr,header=None), yr)
    if '2023' in sheets:
        add('2023', parse_type_c, pd.read_excel(xl,'2023',header=None))
    if '2024' in sheets:
        add('2024', parse_type_d, pd.read_excel(xl,'2024',header=None))
    for yr in ['2025','2026']:
        if yr in sheets:
            add(yr, parse_type_e, pd.read_excel(xl,yr,header=None))

    data['overall'] = compute_overall_from_years(data['years'])
    print("  ✓ Overall totals computed from year sheets")
    return data

# ─── HTML Generator ──────────────────────────────────────────────────────────

def generate_html(data):
    data_json = json.dumps(data, ensure_ascii=False, separators=(',',':'))

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>StartHub Africa — Impact Dashboard</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;600;700;800&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
:root{{
  --bg:#F0F4F9;
  --surface:#FFFFFF;
  --surface2:#F7F9FC;
  --border:#E3E8F0;
  --border2:#C9D4E0;
  --txt:#111827;
  --txt2:#374151;
  --muted:#6B7280;
  --shadow:0 1px 4px rgba(0,0,0,.06),0 1px 2px rgba(0,0,0,.03);
  --shadow-md:0 4px 20px rgba(0,0,0,.09),0 2px 6px rgba(0,0,0,.05);
  --radius:12px;
  /* Brand structural */
  --brand-navy:#243649;
  --brand-blue:#32579F;
  --brand-gold:#DFA12C;
  /* Data colours — full categorical spectrum, max differentiation */
  --c-trained:#3B82F6;
  --c-tot:#8B5CF6;
  --c-biz-sup:#F59E0B;
  --c-biz-sta:#10B981;
  --c-jobs:#EF4444;
  --c-intern:#06B6D4;
  --c-grants-n:#EC4899;
  --c-grants-v:#F97316;
  /* Country colours */
  --uganda:#3B82F6;
  --tanzania:#F59E0B;
  --kenya:#10B981;
}}
html{{font-size:14px;color-scheme:light}}
body{{background:var(--bg);color:var(--txt);font-family:'Plus Jakarta Sans','Segoe UI',system-ui,sans-serif;min-height:100vh}}

/* ── Header ── */
#header{{
  background:var(--brand-navy);border-bottom:3px solid var(--brand-gold);
  padding:12px 24px;display:flex;align-items:center;gap:16px;
  position:sticky;top:0;z-index:200;flex-wrap:wrap;
  box-shadow:0 2px 12px rgba(36,54,73,.35);
}}
.logo{{display:flex;align-items:center;gap:10px;flex:1;min-width:180px}}
.logo svg{{width:36px;height:36px}}
.logo-text h1{{font-size:1rem;font-weight:800;color:#FFFFFF;line-height:1;letter-spacing:-.01em}}
.logo-text span{{font-size:.68rem;color:rgba(255,255,255,.55);letter-spacing:.02em}}
.header-filters{{display:flex;gap:8px;flex-wrap:wrap;align-items:center}}

/* Filter dropdown pills */
.filter-dd{{
  position:relative;display:flex;align-items:center;gap:6px;
  background:rgba(255,255,255,.10);border:1px solid rgba(255,255,255,.20);border-radius:6px;
  padding:5px 10px;cursor:pointer;user-select:none;
  transition:border-color .15s,background .15s;
}}
.filter-dd:hover,.filter-dd.open{{background:rgba(255,255,255,.18);border-color:var(--brand-gold)}}
.filter-dd .fd-label{{font-size:.68rem;font-weight:700;color:rgba(255,255,255,.55);white-space:nowrap;text-transform:uppercase;letter-spacing:.05em}}
.filter-dd .fd-val{{font-size:.78rem;font-weight:700;color:#FFFFFF;max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
.filter-dd .fd-caret{{font-size:.6rem;color:rgba(255,255,255,.45);transition:transform .15s}}
.filter-dd.open .fd-caret{{transform:rotate(180deg)}}

/* Dropdown panel */
.fd-panel{{
  display:none;position:absolute;top:calc(100% + 6px);left:0;
  background:var(--surface);border:1px solid var(--border);border-radius:10px;
  box-shadow:0 8px 24px rgba(0,0,0,.12);
  min-width:200px;max-height:320px;overflow-y:auto;
  z-index:400;padding:6px 0;
}}
.fd-panel.open{{display:block}}
.fd-row{{
  display:flex;align-items:center;gap:9px;
  padding:7px 14px;cursor:pointer;
  transition:background .08s;
}}
.fd-row:hover{{background:var(--surface2)}}
.fd-row.all-row{{border-bottom:1px solid var(--border);font-weight:700}}
.fd-row input[type=checkbox]{{width:14px;height:14px;cursor:pointer;accent-color:var(--brand-blue);flex-shrink:0}}
.fd-row span{{font-size:.74rem;font-weight:600;color:var(--txt2);line-height:1.3}}
.fd-row.all-row span{{color:var(--txt);font-weight:700}}
.fd-badge{{
  display:inline-flex;align-items:center;justify-content:center;
  min-width:18px;height:18px;padding:0 5px;
  background:var(--brand-gold);color:#fff;border-radius:9px;
  font-size:.6rem;font-weight:800;margin-left:4px;flex-shrink:0;
}}
.fd-search-wrap{{padding:8px 10px 6px;border-bottom:1px solid var(--border);background:var(--surface2)}}
.fd-search{{
  width:100%;box-sizing:border-box;padding:5px 8px 5px 26px;
  border:1px solid var(--border2);border-radius:6px;
  font-size:.75rem;background:var(--surface);color:var(--txt);outline:none;
  background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%2394A3B8' stroke-width='2'%3E%3Ccircle cx='11' cy='11' r='8'/%3E%3Cpath d='m21 21-4.35-4.35'/%3E%3C/svg%3E");
  background-repeat:no-repeat;background-position:8px center;
}}
.fd-search:focus{{border-color:var(--brand-blue);background:#EEF3FB}}

select{{background:transparent;border:none;color:var(--txt);font-size:.8rem;font-weight:600;cursor:pointer;outline:none;padding:2px 0}}
.updated{{font-size:.68rem;color:rgba(255,255,255,.45);margin-left:auto;white-space:nowrap}}

/* ── Layout ── */
#main{{padding:22px 28px;max-width:1600px;margin:0 auto}}
.sec-label{{
  font-size:.7rem;font-weight:800;text-transform:uppercase;letter-spacing:.1em;
  color:var(--brand-navy);margin-bottom:14px;
  display:flex;align-items:center;gap:10px;
}}
.sec-label::before{{content:'';width:4px;height:14px;background:var(--brand-gold);border-radius:2px;flex-shrink:0}}
.sec-label::after{{content:'';flex:1;height:1px;background:var(--border)}}
.mb{{margin-bottom:26px}}

/* ── KPI Cards — bento grid ── */
#kpi-grid{{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}}
@media(max-width:1000px){{#kpi-grid{{grid-template-columns:repeat(2,1fr)}}}}
@media(max-width:560px){{#kpi-grid{{grid-template-columns:1fr}}}}
.kpi{{
  background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);
  padding:18px 20px 16px 22px;border-left:4px solid var(--accent,#3B82F6);
  box-shadow:var(--shadow);transition:box-shadow .2s,transform .2s;
  position:relative;overflow:hidden;
}}
.kpi.hero{{grid-column:span 2}}
.kpi::before{{
  content:'';position:absolute;top:0;right:0;width:120px;height:100%;
  background:linear-gradient(to left,var(--accent-bg,transparent),transparent);
  pointer-events:none;
}}
.kpi:hover{{box-shadow:var(--shadow-md);transform:translateY(-2px)}}
.kpi-label{{font-size:.63rem;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:var(--muted);margin-bottom:10px}}
.kpi-value{{font-size:2.1rem;font-weight:800;line-height:1;font-variant-numeric:tabular-nums}}
.kpi.hero .kpi-value{{font-size:2.6rem}}
.kpi-sub{{font-size:.68rem;color:var(--muted);margin-top:8px;display:flex;gap:6px;flex-wrap:wrap}}
.badge{{display:inline-flex;align-items:center;gap:3px;font-size:.65rem;font-weight:600;padding:2px 8px;border-radius:20px}}
.badge-m{{background:#DBEAFE;color:#1E40AF}}
.badge-f{{background:#FCE7F3;color:#9D174D}}

/* ── Chart cards ── */
.chart-card{{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:20px 22px;box-shadow:var(--shadow)}}
.chart-card h3{{font-size:.9rem;font-weight:700;color:var(--txt);margin-bottom:3px}}
.chart-meta{{font-size:.69rem;color:var(--muted);margin-bottom:14px}}
.chart-wrap canvas{{width:100%!important}}
.charts-2{{display:grid;grid-template-columns:2fr 1fr;gap:16px}}
.charts-eq{{display:grid;grid-template-columns:1fr 1fr;gap:16px}}
@media(max-width:900px){{.charts-2,.charts-eq{{grid-template-columns:1fr}}}}

/* ── Metric Checkboxes ── */
.mcheck-group{{display:flex;flex-wrap:wrap;gap:4px 8px;padding:10px 12px;background:var(--surface2);border:1px solid var(--border);border-radius:8px}}
.mcheck-item{{
  display:flex;align-items:center;gap:6px;cursor:pointer;user-select:none;white-space:nowrap;
  padding:4px 10px;border-radius:20px;border:1.5px solid transparent;
  transition:background .12s,border-color .12s;
}}
.mcheck-item:has(input:checked){{background:var(--surface);border-color:var(--border2);box-shadow:var(--shadow)}}
.mcheck-item:hover{{background:var(--surface);border-color:var(--border)}}
.mcheck-item input[type=checkbox]{{display:none}}
.mcheck-dot{{width:10px;height:10px;border-radius:50%;background:var(--chk-col,#3B82F6);flex-shrink:0;opacity:.4;transition:opacity .12s}}
.mcheck-item:has(input:checked) .mcheck-dot{{opacity:1}}
.mcheck-item span{{font-size:.72rem;font-weight:600;color:var(--muted)}}
.mcheck-item:has(input:checked) span{{color:var(--txt2)}}
.mcheck-item.all-item{{background:var(--surface);border-color:var(--border2)}}
.mcheck-item.all-item span{{font-weight:700;color:var(--txt)}}
.mcheck-sep{{display:none}}

/* ── Country Metrics Table ── */
.country-metrics-table{{width:100%;border-collapse:collapse;font-size:.76rem}}
.country-metrics-table th{{
  padding:8px 12px;text-align:left;font-weight:700;font-size:.65rem;
  text-transform:uppercase;letter-spacing:.06em;color:var(--muted);
  border-bottom:2px solid var(--border);white-space:nowrap;
}}
.country-metrics-table td{{
  padding:9px 12px;border-bottom:1px solid var(--border);white-space:nowrap;
}}
.country-metrics-table tr:last-child td{{border-bottom:none}}
.country-metrics-table tr:hover td{{background:var(--surface2)}}
.c-badge{{
  display:inline-block;font-size:.6rem;font-weight:800;padding:2px 6px;
  border-radius:4px;color:#fff;margin-right:6px;letter-spacing:.04em;
}}
.metric-bar-cell{{display:flex;align-items:center;gap:6px}}
.mini-bar{{height:5px;border-radius:3px;min-width:2px;flex-shrink:0}}
.num-cell{{font-variant-numeric:tabular-nums;font-weight:700;color:var(--txt);text-align:right}}
.country-col-head{{display:flex;align-items:center;gap:8px;font-size:.78rem;font-weight:700}}

/* ── Project Drill-Down ── */
.drill-controls{{display:flex;gap:10px;margin-bottom:12px;align-items:center;flex-wrap:wrap}}
.drill-controls label{{font-size:.7rem;font-weight:600;color:var(--muted)}}
.ctrl-select{{
  background:var(--surface);border:1px solid var(--border2);color:var(--txt);
  border-radius:6px;padding:5px 10px;font-size:.78rem;font-weight:600;
  cursor:pointer;outline:none;
}}
.ctrl-select:focus{{border-color:var(--brand-blue)}}

/* Project search + selector */
.drill-selector{{
  display:flex;gap:12px;margin-bottom:14px;align-items:flex-start;
}}
.drill-selector-panel{{
  display:flex;flex-direction:column;gap:0;
  width:280px;flex-shrink:0;
  border:1px solid var(--border);border-radius:8px;
  overflow:hidden;background:var(--surface);
}}
.drill-search{{
  width:100%;box-sizing:border-box;padding:8px 10px 8px 32px;
  border:none;border-bottom:1px solid var(--border);
  font-size:.78rem;background:var(--surface2);color:var(--txt);outline:none;
  background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='14' height='14' viewBox='0 0 24 24' fill='none' stroke='%2394A3B8' stroke-width='2'%3E%3Ccircle cx='11' cy='11' r='8'/%3E%3Cpath d='m21 21-4.35-4.35'/%3E%3C/svg%3E");
  background-repeat:no-repeat;background-position:10px center;
  border-radius:0;
}}
.drill-search:focus{{background:#EEF3FB}}
.drill-proj-box{{
  display:flex;flex-direction:column;gap:0;
  overflow-y:auto;max-height:220px;
  padding:4px 0;
}}
.dp-item{{
  display:flex;align-items:center;gap:8px;cursor:pointer;
  user-select:none;padding:6px 12px;
  transition:background .1s;
}}
.dp-item:hover{{background:var(--surface2)}}
.dp-item.all-item{{background:var(--surface2);border-bottom:1px solid var(--border);padding:7px 12px}}
.dp-item.all-item:hover{{background:#E8EDF5}}
.dp-item input[type=checkbox]{{width:14px;height:14px;cursor:pointer;accent-color:var(--brand-blue);flex-shrink:0}}
.dp-item span{{font-size:.72rem;font-weight:600;color:var(--txt2);line-height:1.3}}
.dp-item.all-item span{{font-weight:700;color:var(--txt)}}
.dp-sep{{display:none}}
.dp-none{{font-size:.72rem;color:var(--muted);padding:10px 12px}}
.drill-hint{{
  font-size:.68rem;color:var(--muted);padding:4px 12px 6px;
  border-top:1px solid var(--border);background:var(--surface2);
  text-align:center;
}}
@media(max-width:700px){{.drill-selector{{flex-direction:column}}.drill-selector-panel{{width:100%}}}}

.drill-outer{{overflow-x:auto;border:1px solid var(--border);border-radius:var(--radius);background:var(--surface)}}
.drill-table{{border-collapse:separate;border-spacing:0;table-layout:auto;min-width:100%}}

/* Sticky first column */
.drill-table .metric-col{{
  position:sticky;left:0;z-index:10;background:var(--surface);
  min-width:170px;max-width:170px;
  border-right:2px solid var(--border);
  padding:8px 12px;font-size:.72rem;color:var(--txt2);white-space:nowrap;
}}
.drill-table thead .metric-col{{
  background:var(--surface2);font-weight:800;font-size:.65rem;text-transform:uppercase;
  letter-spacing:.07em;color:var(--muted);border-bottom:2px solid var(--border);
  position:sticky;left:0;top:0;z-index:20;
}}
.drill-table thead th{{
  background:var(--surface2);border-bottom:2px solid var(--border);
  padding:8px 12px;min-width:140px;max-width:180px;
  vertical-align:bottom;
}}
.proj-head{{
  display:flex;flex-direction:column;gap:4px;
  font-size:.7rem;font-weight:700;color:var(--txt);
}}
.proj-head .proj-name{{
  overflow:hidden;text-overflow:ellipsis;white-space:nowrap;
  max-width:150px;font-size:.72rem;
}}
.proj-head .proj-flag{{
  display:inline-flex;align-items:center;gap:4px;font-size:.6rem;font-weight:700;
  padding:2px 5px;border-radius:3px;color:#fff;width:fit-content;
}}
.drill-table tbody tr:hover .metric-col,
.drill-table tbody tr:hover td{{background:#F4F7FB}}
.drill-table tbody tr.sort-row .metric-col{{background:#EBF0F7}}
.drill-table tbody tr.sort-row td{{background:#EBF0F7}}
.drill-table tbody td{{
  padding:7px 12px;border-bottom:1px solid var(--border);vertical-align:middle;
}}
.cell-inner{{display:flex;flex-direction:column;gap:3px}}
.cell-num{{font-size:.78rem;font-weight:700;color:var(--txt);font-variant-numeric:tabular-nums;white-space:nowrap}}
.cell-num.zero{{color:var(--muted);font-weight:400}}
.cell-bar-wrap{{height:4px;background:var(--border);border-radius:2px;overflow:hidden}}
.cell-bar-fill{{height:100%;border-radius:2px;transition:width .3s ease}}

/* ── Misc ── */
.no-data{{text-align:center;color:var(--muted);padding:28px;font-size:.82rem}}
#modal-overlay{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.4);z-index:500;align-items:center;justify-content:center}}
#modal-overlay.open{{display:flex}}
#modal{{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:24px;width:min(820px,95vw);max-height:90vh;overflow-y:auto;position:relative;box-shadow:var(--shadow-md)}}
#modal h2{{font-size:1rem;margin-bottom:2px;color:var(--txt)}}
#modal .modal-sub{{font-size:.72rem;color:var(--muted);margin-bottom:16px}}
#modal-close{{position:absolute;top:14px;right:16px;background:none;border:none;color:var(--muted);cursor:pointer;font-size:1.1rem}}
::-webkit-scrollbar{{width:5px;height:5px}}
::-webkit-scrollbar-track{{background:var(--surface2)}}
::-webkit-scrollbar-thumb{{background:var(--border2);border-radius:3px}}
</style>
</head>
<body>

<!-- Header -->
<div id="header">
  <div class="logo">
    <svg viewBox="0 0 36 36" fill="none" xmlns="http://www.w3.org/2000/svg">
      <rect width="36" height="36" rx="8" fill="#1A2A38"/>
      <!-- Bulb body -->
      <ellipse cx="18" cy="15" rx="7" ry="8" fill="none" stroke="#DFA12C" stroke-width="1.8"/>
      <!-- Person inside bulb -->
      <circle cx="18" cy="12" r="2" fill="#32579F"/>
      <path d="M15 17 Q18 20 21 17" stroke="#32579F" stroke-width="1.5" stroke-linecap="round" fill="none"/>
      <!-- Filament lines -->
      <path d="M15.5 23h5M16 25h4" stroke="#DFA12C" stroke-width="1.5" stroke-linecap="round"/>
      <!-- Left leaf -->
      <path d="M11 22 Q7 19 10 15" stroke="#2D6B2A" stroke-width="1.8" stroke-linecap="round" fill="none"/>
      <!-- Right leaf -->
      <path d="M25 22 Q29 19 26 15" stroke="#2D6B2A" stroke-width="1.8" stroke-linecap="round" fill="none"/>
    </svg>
    <div class="logo-text"><h1>StartHub Africa</h1><span>Impact Dashboard</span></div>
  </div>
  <div class="header-filters">
    <div class="filter-dd" id="fd-year">
      <span class="fd-label">Year</span>
      <span class="fd-val" id="fdv-year">All Years</span>
      <span class="fd-caret">▾</span>
      <div class="fd-panel" id="fdp-year"></div>
    </div>
    <div class="filter-dd" id="fd-country">
      <span class="fd-label">Country</span>
      <span class="fd-val" id="fdv-country">All Countries</span>
      <span class="fd-caret">▾</span>
      <div class="fd-panel" id="fdp-country"></div>
    </div>
    <div class="filter-dd" id="fd-project">
      <span class="fd-label">Project</span>
      <span class="fd-val" id="fdv-project">All Projects</span>
      <span class="fd-caret">▾</span>
      <div class="fd-panel" id="fdp-project"></div>
    </div>
  </div>
  <div class="updated">Updated: <span id="updated-date"></span></div>
</div>

<div id="main">

  <!-- KPI Cards -->
  <div class="sec-label mb" style="margin-bottom:8px">Key Impact Metrics</div>
  <div id="kpi-grid" class="mb"></div>

  <!-- Trend Chart -->
  <div class="sec-label mb" style="margin-bottom:8px">Impact Over Time</div>
  <div class="chart-card mb">
    <h3>Year-on-Year Trends</h3>
    <div class="chart-meta">Select metrics to compare · Click a year bar to see monthly breakdown</div>
    <div class="mcheck-group" id="trend-checks"></div>
    <div class="chart-wrap" style="height:280px;margin-top:12px"><canvas id="trend-chart"></canvas></div>
  </div>

  <!-- Country breakdown -->
  <div class="sec-label mb" style="margin-bottom:8px">Country Breakdown</div>
  <div class="charts-2 mb">
    <div class="chart-card">
      <h3>Metrics by Country</h3>
      <div class="chart-meta">Select any metrics to compare across Uganda · Tanzania · Kenya</div>
      <div class="mcheck-group" id="country-checks"></div>
      <div id="country-table-wrap" style="margin-top:12px"></div>
    </div>
    <div class="chart-card">
      <h3>Training Share</h3>
      <div class="chart-meta">Entrepreneurs trained by country</div>
      <div class="chart-wrap" style="height:260px"><canvas id="donut-chart"></canvas></div>
    </div>
  </div>

  <!-- Gender Split -->
  <div class="sec-label mb" style="margin-bottom:8px">Gender Split — Trainings Only</div>
  <div class="chart-card mb">
    <h3>Male vs Female Trained by Year</h3>
    <div class="chart-meta">Gender breakdown applies to entrepreneurship trainings only</div>
    <div class="chart-wrap" style="height:240px"><canvas id="gender-chart"></canvas></div>
  </div>

  <!-- Grants Panel -->
  <div class="sec-label mb" style="margin-bottom:8px">Grants &amp; Funding</div>
  <div class="chart-card mb">
    <h3>Grant Value by Project (USD)</h3>
    <div class="chart-meta">Projects that received grant funding — filtered by selected period &amp; country</div>
    <div id="grants-wrap" class="chart-wrap" style="height:220px"><canvas id="grants-chart"></canvas></div>
  </div>

  <!-- Project Drill-Down -->
  <div class="sec-label mb" style="margin-bottom:8px">Project Drill-Down</div>
  <div class="chart-card mb">
    <div class="drill-controls">
      <div><label>Sort by &nbsp;</label><select class="ctrl-select" id="drill-sort">
        <option value="trained_total">Trained</option>
        <option value="tot_total">Training of Trainers</option>
        <option value="biz_supported">Biz Supported</option>
        <option value="biz_started">Biz Started</option>
        <option value="jobs">Jobs Placed</option>
        <option value="internships">Internships</option>
        <option value="grants_count">Grants (count)</option>
        <option value="grants_value">Grants (USD)</option>
      </select></div>
      <div><label>Country &nbsp;</label><select class="ctrl-select" id="drill-country">
        <option value="all">All Countries</option>
        <option value="Uganda">Uganda</option>
        <option value="Tanzania">Tanzania</option>
        <option value="Kenya">Kenya</option>
      </select></div>
      <span id="drill-count" style="font-size:.7rem;color:var(--muted);margin-left:4px"></span>
    </div>
    <div class="drill-selector">
      <div class="drill-selector-panel">
        <input type="search" id="drill-search" class="drill-search" placeholder="Search projects…" autocomplete="off">
        <div id="drill-proj-box" class="drill-proj-box"></div>
        <div class="drill-hint" id="drill-hint"></div>
      </div>
      <div id="drill-wrap" style="flex:1;min-width:0;overflow-x:auto"></div>
    </div>
  </div>

</div>

<!-- Monthly Modal -->
<div id="modal-overlay">
  <div id="modal">
    <button id="modal-close">✕</button>
    <h2 id="modal-title"></h2>
    <div class="modal-sub" id="modal-sub"></div>
    <div class="chart-wrap" style="height:280px"><canvas id="modal-chart"></canvas></div>
  </div>
</div>

<script>
const DATA = {data_json};

// ── Config ─────────────────────────────────────────────────────────────────
// Categorical palette — maximally distinct hues, ordered by importance for bento layout
// Row 1: Trained(hero×2) | Biz Supported | Biz Started
// Row 2: Jobs | Internships | Grants USD(hero×2)
// Row 3: Training of Trainers(hero×2) | Grants Count(hero×2)
const METRICS = [
  {{key:'trained_total', label:'Entrepreneurs Trained', icon:'🎓', color:'#3B82F6', prefix:'',  hero:true}},
  {{key:'biz_supported', label:'Biz Supported',         icon:'🏢', color:'#F59E0B', prefix:'',  hero:false}},
  {{key:'biz_started',   label:'Biz Started',           icon:'🚀', color:'#10B981', prefix:'',  hero:false}},
  {{key:'jobs',          label:'Jobs Placed',            icon:'💼', color:'#EF4444', prefix:'',  hero:false}},
  {{key:'internships',   label:'Internships',            icon:'📋', color:'#06B6D4', prefix:'',  hero:false}},
  {{key:'grants_value',  label:'Grants (USD)',           icon:'💰', color:'#F97316', prefix:'$', hero:true}},
  {{key:'tot_total',     label:'Training of Trainers',  icon:'📚', color:'#8B5CF6', prefix:'',  hero:true}},
  {{key:'grants_count',  label:'Grants Awarded',        icon:'🏷️', color:'#EC4899', prefix:'',  hero:true}},
];
const METRIC_MAP = Object.fromEntries(METRICS.map(m=>[m.key,m]));

// Country colours — vivid, distinct
const COUNTRY_COLORS = {{Uganda:'#3B82F6', Tanzania:'#F59E0B', Kenya:'#10B981'}};
const COUNTRY_BG     = {{Uganda:'#EFF6FF', Tanzania:'#FFFBEB', Kenya:'#ECFDF5'}};
const COUNTRIES = ['Uganda','Tanzania','Kenya'];

// Categorical pastel backgrounds and borders for project columns
const PROJ_BG = [
  '#EFF6FF','#F5F3FF','#FFFBEB','#ECFDF5','#FEF2F2',
  '#ECFEFF','#FDF2F8','#FFF7ED','#F0F9FF','#F0FDF4',
  '#EEF2FF','#FFF1F2','#FEFCE8','#F0FDFA','#FAF5FF',
  '#FFF8F0','#F0F4FF','#FFFAF0','#F5FFFA','#FFF0F5',
];
const PROJ_BORDER = [
  '#3B82F6','#8B5CF6','#F59E0B','#10B981','#EF4444',
  '#06B6D4','#EC4899','#F97316','#6366F1','#22C55E',
  '#818CF8','#FB7185','#FCD34D','#34D399','#60A5FA',
  '#FB923C','#A78BFA','#FCA5A5','#6EE7B7','#93C5FD',
];

// ── State ──────────────────────────────────────────────────────────────────
const state = {{
  // null = all selected; Set = specific selections
  activeYears: null,
  activeCountries: null,
  activeProjects: null,
  activeMetrics: METRICS.map(m=>m.key),
  countryMetrics: METRICS.map(m=>m.key),
  drillSort:'trained_total', drillCountry:'all',
  drillSearch:'', drillSelected:null,
}};

let charts = {{}};
let modalChart = null;

// ── Data helpers ───────────────────────────────────────────────────────────
const getYears = () => Object.keys(DATA.years).sort();

// Active filter getters (null state = everything)
const getActiveYears     = () => state.activeYears     ? [...state.activeYears].sort()     : getYears();
const getActiveCountries = () => state.activeCountries ? [...state.activeCountries]         : [...COUNTRIES];

// Full blank covering all raw metric keys
function blankFull(){{
  return {{trained_m:0,trained_f:0,trained_total:0,tot_m:0,tot_f:0,tot_total:0,
           biz_supported:0,biz_started:0,jobs:0,internships:0,
           grants_count:0,grants_value:0,revenue_baseline:0,revenue_endline:0,cohort_companies:0}};
}}

// Core aggregator — sums metrics for given years + countries, respecting project filter
function aggregate(years, countries){{
  const yrs  = years     || getYears();
  const ctrs = countries || COUNTRIES;
  const tot  = blankFull();
  for(const yr of yrs){{
    const yd = DATA.years[yr]; if(!yd) continue;
    for(const c of ctrs){{
      if(state.activeProjects){{
        // sum only matching project rows
        for(const p of (yd.countries?.[c]?.projects||[])){{
          if(state.activeProjects.has(normName(p.name).toLowerCase())){{
            Object.keys(p.metrics||{{}}).forEach(k=>{{ tot[k]=(tot[k]||0)+(p.metrics[k]||0); }});
          }}
        }}
      }} else {{
        const m = yd.countries?.[c]?.totals;
        if(m) Object.keys(tot).forEach(k=>{{ tot[k]=(tot[k]||0)+(m[k]||0); }});
      }}
    }}
  }}
  return tot;
}}

function getFiltered(){{
  return aggregate(getActiveYears(), getActiveCountries());
}}

// Legacy shim used by modal and any remaining callers
function getMetrics(yr, country){{
  return aggregate([yr], country==='all' ? null : [country]);
}}

// Project names for a specific set of years (drives the project dropdown)
function getProjectNamesForYears(years){{
  const seen = new Set(); const names = [];
  for(const yr of years){{
    const yd = DATA.years[yr]; if(!yd) continue;
    for(const c of COUNTRIES){{
      for(const p of (yd.countries?.[c]?.projects||[])){{
        const key = normName(p.name).toLowerCase();
        if(!seen.has(key)){{
          seen.add(key);
          names.push({{key, label:normName(p.name).replace(/\b\w/g,ch=>ch.toUpperCase())}});
        }}
      }}
    }}
  }}
  return names.sort((a,b)=>a.label.localeCompare(b.label));
}}

// All unique project names regardless of year
function getAllProjectNames(){{ return getProjectNamesForYears(getYears()); }}

function fmt(v, prefix=''){{
  if(!v||v===0) return '—';
  const n = Math.round(v);
  if(n>=1000000) return prefix+(n/1000000).toFixed(1)+'M';
  if(n>=10000)   return prefix+(n/1000).toFixed(0)+'K';
  if(n>=1000)    return prefix+(n/1000).toFixed(1)+'K';
  return prefix+n.toLocaleString();
}}
function fmtFull(v, prefix=''){{
  if(!v||v===0) return '—';
  return prefix+Math.round(v).toLocaleString();
}}
function pct(a,b){{ return b>0?Math.round(a/b*100):0; }}

function normName(n){{ return (n||'').trim().replace(/\s+/g,' '); }}

function getProjects(drillCountry){{
  const years      = getActiveYears();
  const headerCtrs = getActiveCountries();
  const map = {{}};
  for(const yr of years){{
    const yd = DATA.years[yr]; if(!yd) continue;
    for(const c of COUNTRIES){{
      if(drillCountry!=='all'&&c!==drillCountry) continue;
      if(!headerCtrs.includes(c)) continue;
      for(const p of (yd.countries?.[c]?.projects||[])){{
        // Apply programme filter (state.activeProjects) if set
        if(state.activeProjects){{
          if(!state.activeProjects.has(normName(p.name).toLowerCase())) continue;
        }}
        // normalise key: lowercase + trim → merges capitalisation variants
        const key = normName(p.name).toLowerCase()+'||'+c;
        if(map[key]){{
          Object.keys(p.metrics||{{}}).forEach(k=>{{
            map[key].metrics[k]=(map[key].metrics[k]||0)+(p.metrics[k]||0);
          }});
        }} else {{
          const titled = normName(p.name).replace(/\b\w/g,ch=>ch.toUpperCase());
          map[key]={{name:titled,country:c,metrics:{{...(p.metrics||{{}})}}}};
        }}
      }}
    }}
  }}
  return Object.values(map);
}}

// ── Count-up animation ────────────────────────────────────────────────────
function countUp(el, target, prefix, duration){{
  const start = performance.now();
  function step(now){{
    const p = Math.min((now-start)/duration,1);
    const ease = 1-Math.pow(1-p,3); // ease-out cubic
    const v = Math.round(target*ease);
    el.textContent = prefix + v.toLocaleString();
    if(p<1) requestAnimationFrame(step);
    else el.textContent = prefix + target.toLocaleString();
  }}
  requestAnimationFrame(step);
}}

// ── KPI Cards ─────────────────────────────────────────────────────────────
function renderKPI(){{
  const m = getFiltered();
  const rows = METRICS.map(met=>{{
    const v = Math.round(m[met.key]||0);
    let sub = '';
    if(met.key==='trained_total'){{
      const tm=m.trained_m||0, tf=m.trained_f||0;
      if(tm||tf) sub=`<span class="badge badge-m">♂ ${{Math.round(tm).toLocaleString()}}</span>`+
                     `<span class="badge badge-f">♀ ${{Math.round(tf).toLocaleString()}}</span>`;
    }}
    if(met.key==='tot_total'){{
      const tm=m.tot_m||0, tf=m.tot_f||0;
      if(tm||tf) sub=`<span class="badge badge-m">♂ ${{Math.round(tm).toLocaleString()}}</span>`+
                     `<span class="badge badge-f">♀ ${{Math.round(tf).toLocaleString()}}</span>`;
    }}
    if(met.key==='grants_value'){{
      const gc=m.grants_count||0;
      if(gc) sub=`<span style="color:var(--muted)">${{Math.round(gc).toLocaleString()}} grants</span>`;
    }}
    const heroClass = met.hero?'hero':'';
    return `<div class="kpi ${{heroClass}}" style="--accent:${{met.color}};--accent-bg:${{met.color}}18" data-kpi-val="${{v}}" data-kpi-prefix="${{met.prefix}}">
      <div class="kpi-label">${{met.label}}</div>
      <div class="kpi-value" style="color:${{met.color}}">${{met.prefix}}${{v.toLocaleString()}}</div>
      ${{sub?`<div class="kpi-sub">${{sub}}</div>`:''}}
    </div>`;
  }});
  document.getElementById('kpi-grid').innerHTML = rows.join('');
  // Trigger count-up on each card
  document.querySelectorAll('#kpi-grid .kpi').forEach(card=>{{
    const val = parseFloat(card.dataset.kpiVal)||0;
    const pfx = card.dataset.kpiPrefix||'';
    const el  = card.querySelector('.kpi-value');
    if(el && val>0) countUp(el, val, pfx, 900);
  }});
}}

// ── Checkbox group builder (reusable) ────────────────────────────────────────
function buildCheckGroup(containerId, stateKey, onChangeFn){{
  const wrap = document.getElementById(containerId);
  const allChecked = () => state[stateKey].length === METRICS.length;

  wrap.innerHTML = `
    <label class="mcheck-item all-item">
      <input type="checkbox" id="${{containerId}}-all" ${{allChecked()?'checked':''}}>
      <span>All Metrics</span>
    </label>
    ${{METRICS.map(m=>`
      <label class="mcheck-item" style="--chk-col:${{m.color}}">
        <input type="checkbox" data-mk="${{m.key}}" ${{state[stateKey].includes(m.key)?'checked':''}}>
        <span class="mcheck-dot"></span>
        <span>${{m.label}}</span>
      </label>`).join('')}}`;

  // "Select All" checkbox
  const allBox = wrap.querySelector(`#${{containerId}}-all`);
  allBox.addEventListener('change', ()=>{{
    state[stateKey] = allBox.checked ? METRICS.map(m=>m.key) : [];
    buildCheckGroup(containerId, stateKey, onChangeFn);
    onChangeFn();
  }});

  // Individual checkboxes
  wrap.querySelectorAll('input[data-mk]').forEach(inp=>{{
    inp.addEventListener('change', ()=>{{
      if(inp.checked){{
        if(!state[stateKey].includes(inp.dataset.mk)) state[stateKey].push(inp.dataset.mk);
      }} else {{
        state[stateKey] = state[stateKey].filter(k=>k!==inp.dataset.mk);
      }}
      // Sync "Select All" state
      allBox.checked = state[stateKey].length === METRICS.length;
      allBox.indeterminate = state[stateKey].length > 0 && state[stateKey].length < METRICS.length;
      onChangeFn();
    }});
  }});

  // Set indeterminate on load if partial
  allBox.indeterminate = state[stateKey].length > 0 && state[stateKey].length < METRICS.length;
}}

// ── Chart.js global defaults ────────────────────────────────────────────────
Chart.defaults.font.family = "'Plus Jakarta Sans','Segoe UI',system-ui,sans-serif";
Chart.defaults.font.size   = 12;
Chart.defaults.animation   = {{ duration:900, easing:'easeInOutQuart' }};
Chart.defaults.plugins.legend.labels.boxWidth = 12;
Chart.defaults.plugins.legend.labels.padding  = 16;

// ── Trend Chart ─────────────────────────────────────────────────────────────
function renderTrendChecks(){{
  buildCheckGroup('trend-checks', 'activeMetrics', updateTrend);
}}

function buildTrendDatasets(){{
  const years = getActiveYears();
  const ctrs  = getActiveCountries();
  return {{
    labels: years,
    datasets: METRICS.filter(m=>state.activeMetrics.includes(m.key)).map(m=>{{
      const vals = years.map(yr=>aggregate([yr],ctrs)[m.key]||0);
      return {{
        label:m.label, data:vals,
        borderColor:m.color, backgroundColor:m.color+'28',
        borderWidth:3, pointRadius:5, pointHoverRadius:9,
        pointBackgroundColor:m.color, pointBorderColor:'#fff', pointBorderWidth:2,
        fill:false, tension:0.4,
        yAxisID: m.key==='grants_value'?'y2':'y1'
      }};
    }})
  }};
}}

function initTrend(){{
  const ctx = document.getElementById('trend-chart');
  charts.trend = new Chart(ctx,{{
    type:'line', data:buildTrendDatasets(),
    options:{{
      responsive:true, maintainAspectRatio:false,
      interaction:{{mode:'index',intersect:false}},
      onClick:(e,els)=>{{ if(els.length) openModal(getYears()[els[0].index]); }},
      plugins:{{
        legend:{{display:false}},
        tooltip:{{
          backgroundColor:'#243649',borderColor:'#DFA12C',borderWidth:1,
          titleColor:'#F1F5F9',bodyColor:'#94A3B8',padding:10,
          callbacks:{{label:c=>` ${{c.dataset.label}}: ${{c.parsed.y?.toLocaleString?.()??c.parsed.y}}`}}
        }}
      }},
      scales:{{
        x:{{grid:{{color:'#F0F4F9'}},ticks:{{color:'#6B7280',font:{{size:11}}}}}},
        y1:{{position:'left',grid:{{color:'#F0F4F9'}},ticks:{{color:'#6B7280',font:{{size:11}},callback:v=>v>=1000000?(v/1000000).toFixed(1)+'M':v>=1000?(v/1000).toFixed(0)+'K':v}}}},
        y2:{{position:'right',grid:{{drawOnChartArea:false}},ticks:{{color:'#6B7280',font:{{size:11}},callback:v=>'$'+(v>=1000?(v/1000).toFixed(0)+'K':v)}}}}
      }}
    }}
  }});
}}

function updateTrend(){{
  if(!charts.trend) return;
  charts.trend.data = buildTrendDatasets();
  charts.trend.update('active');
}}

// ── Donut Chart ─────────────────────────────────────────────────────────────
function renderDonut(){{
  const yrs  = getActiveYears();
  const ctrs = getActiveCountries();
  const vals = ctrs.map(c=>aggregate(yrs,[c]).trained_total||0);
  const ctx = document.getElementById('donut-chart');
  if(charts.donut) charts.donut.destroy();
  charts.donut = new Chart(ctx,{{
    type:'doughnut',
    data:{{labels:ctrs,datasets:[{{
      data:vals, backgroundColor:ctrs.map(c=>COUNTRY_BG[c]),
      borderColor:ctrs.map(c=>COUNTRY_COLORS[c]), borderWidth:2, hoverOffset:8
    }}]}},
    options:{{
      responsive:true,maintainAspectRatio:false,cutout:'62%',
      plugins:{{
        legend:{{position:'bottom',labels:{{color:'#334155',font:{{size:11,family:"'Plus Jakarta Sans',sans-serif"}},padding:14,usePointStyle:true}}}},
        tooltip:{{
          backgroundColor:'#1E293B',borderColor:'#334155',borderWidth:1,
          callbacks:{{label:c=>`${{c.label}}: ${{c.parsed.toLocaleString()}} (${{pct(c.parsed,vals.reduce((a,b)=>a+b,0))}}%)`}}
        }}
      }}
    }}
  }});
}}

// ── Country Metrics Checks ────────────────────────────────────────────────────
function renderCountryChecks(){{
  buildCheckGroup('country-checks', 'countryMetrics', renderCountryTable);
}}

// ── Country Metrics Table ────────────────────────────────────────────────────
function renderCountryTable(){{
  const yrs  = getActiveYears();
  const ctrs = getActiveCountries();
  const cData = {{}};
  for(const c of ctrs) cData[c] = aggregate(yrs,[c]);

  const visibleMetrics = METRICS.filter(m=>state.countryMetrics.includes(m.key));
  const maxes = {{}};
  visibleMetrics.forEach(met=>{{
    maxes[met.key]=Math.max(...ctrs.map(c=>cData[c][met.key]||0),1);
  }});

  let html=`<div style="overflow-x:auto"><table class="country-metrics-table">
    <thead><tr>
      <th style="min-width:160px">Metric</th>
      ${{ctrs.map(c=>`<th style="min-width:140px"><div class="country-col-head">
        <span class="c-badge" style="background:${{COUNTRY_COLORS[c]}}">${{c.substring(0,2).toUpperCase()}}</span>${{c}}
      </div></th>`).join('')}}
    </tr></thead>
    <tbody>`;

  for(const met of visibleMetrics){{
    html+=`<tr><td class="metric-col" style="position:unset;border-right:none;max-width:none">${{met.icon}} ${{met.label}}</td>`;
    for(const c of ctrs){{
      const v=cData[c][met.key]||0;
      const barW=Math.round(v/maxes[met.key]*100);
      html+=`<td>
        <div class="metric-bar-cell">
          <span class="num-cell" style="min-width:60px">${{fmtFull(v,met.prefix)}}</span>
          ${{v>0?`<div style="flex:1;background:var(--border);border-radius:2px;height:5px">
            <div style="width:${{barW}}%;height:100%;background:${{COUNTRY_COLORS[c]}};border-radius:2px"></div>
          </div>`:''}}
        </div>
      </td>`;
    }}
    html+=`</tr>`;
  }}
  html+=`</tbody></table></div>`;
  document.getElementById('country-table-wrap').innerHTML = html;
}}

function blank(){{ return Object.fromEntries(METRICS.map(m=>[m.key,0])); }}

// ── Gender Chart ─────────────────────────────────────────────────────────────
function renderGender(){{
  const years = getActiveYears();
  const ctrs  = getActiveCountries();
  const mVals = years.map(yr=>aggregate([yr],ctrs).trained_m||0);
  const fVals = years.map(yr=>aggregate([yr],ctrs).trained_f||0);
  const ctx = document.getElementById('gender-chart');
  if(charts.gender) charts.gender.destroy();
  charts.gender = new Chart(ctx,{{
    type:'bar',
    data:{{labels:years,datasets:[
      {{label:'Male',data:mVals,backgroundColor:'#BFDBFE',borderColor:'#3B82F6',borderWidth:1.5,borderRadius:4,stack:'g'}},
      {{label:'Female',data:fVals,backgroundColor:'#FBCFE8',borderColor:'#EC4899',borderWidth:1.5,borderRadius:4,stack:'g'}},
    ]}},
    options:{{
      responsive:true,maintainAspectRatio:false,
      plugins:{{
        legend:{{position:'top',labels:{{color:'#334155',font:{{size:11,family:"'Plus Jakarta Sans',sans-serif"}},usePointStyle:true}}}},
        tooltip:{{backgroundColor:'#243649',borderColor:'#DFA12C',borderWidth:1,
          callbacks:{{afterBody:items=>{{const t=items.reduce((s,i)=>s+i.parsed.y,0);return t?[`Total: ${{t.toLocaleString()}}`]:[];}}}}
        }}
      }},
      scales:{{
        x:{{stacked:true,grid:{{color:'#F0F4F9'}},ticks:{{color:'#6B7280',font:{{size:11}}}}}},
        y:{{stacked:true,grid:{{color:'#F0F4F9'}},ticks:{{color:'#6B7280',font:{{size:11}},callback:v=>v>=1000?(v/1000).toFixed(0)+'K':v}}}}
      }}
    }}
  }});
}}

// ── Grants Chart ─────────────────────────────────────────────────────────────
function renderGrants(){{
  const ps = getProjects('all').filter(p=>(p.metrics.grants_value||0)>0)
    .sort((a,b)=>(b.metrics.grants_value||0)-(a.metrics.grants_value||0)).slice(0,14);
  const wrap = document.getElementById('grants-wrap');
  if(!ps.length){{ wrap.innerHTML='<div class="no-data">No grant data for selected filters</div>'; if(charts.grants)charts.grants.destroy(); return; }}
  if(!wrap.querySelector('canvas')){{ wrap.innerHTML='<canvas id="grants-chart"></canvas>'; }}
  const ctx = document.getElementById('grants-chart');
  if(charts.grants) charts.grants.destroy();
  charts.grants = new Chart(ctx,{{
    type:'bar',
    data:{{
      labels:ps.map(p=>p.name.length>35?p.name.slice(0,35)+'…':p.name),
      datasets:[{{
        label:'Grant Value (USD)',
        data:ps.map(p=>p.metrics.grants_value||0),
        backgroundColor:ps.map(p=>COUNTRY_BG[p.country]),
        borderColor:ps.map(p=>COUNTRY_COLORS[p.country]),
        borderWidth:1.5,borderRadius:4
      }}]
    }},
    options:{{
      indexAxis:'y',responsive:true,maintainAspectRatio:false,
      plugins:{{
        legend:{{display:false}},
        tooltip:{{backgroundColor:'#243649',borderColor:'#DFA12C',borderWidth:1,
          callbacks:{{label:c=>`USD ${{Math.round(c.parsed.x).toLocaleString()}} · ${{ps[c.dataIndex].country}}`}}
        }}
      }},
      scales:{{
        x:{{grid:{{color:'#F0F4F9'}},ticks:{{color:'#6B7280',font:{{size:11}},callback:v=>'$'+v.toLocaleString()}}}},
        y:{{grid:{{display:false}},ticks:{{color:'#6B7280',font:{{size:10}}}}}}
      }}
    }}
  }});
}}

// ── Project Selector ──────────────────────────────────────────────────────────
function renderDrillSelector(projects){{
  const q = state.drillSearch.toLowerCase().trim();
  const visible = q
    ? projects.filter(p=>p.name.toLowerCase().includes(q)||p.country.toLowerCase().includes(q))
    : projects;
  const allKeys = projects.map(p=>p.name+'||'+p.country);
  const sel = state.drillSelected || new Set(allKeys);
  const box = document.getElementById('drill-proj-box');
  const hint = document.getElementById('drill-hint');
  if(!box) return;

  if(!visible.length){{
    box.innerHTML = '<span class="dp-none">No projects match.</span>';
    if(hint) hint.textContent='';
    return;
  }}

  const visKeys = visible.map(p=>p.name+'||'+p.country);
  const selCount = visKeys.filter(k=>sel.has(k)).length;
  const allChecked = selCount===visKeys.length;
  const noneChecked = selCount===0;

  // Build rows
  let h = `<label class="dp-item all-item"><input type="checkbox" id="dp-all"${{allChecked?' checked':''}}>`+
          `<span>Select All (${{visible.length}})</span></label>`;
  for(const p of visible){{
    const k = p.name+'||'+p.country;
    const checked = sel.has(k);
    const flag = p.country.substring(0,2).toUpperCase();
    h+=`<label class="dp-item"><input type="checkbox" class="dp-cb" data-key="${{k}}"${{checked?' checked':''}}>`+
       `<span title="${{p.name}}">${{p.name}}<em style="color:var(--muted);font-style:normal;font-size:.65rem;margin-left:4px">[${{flag}}]</em></span></label>`;
  }}
  box.innerHTML = h;

  // Hint counter
  const totalSel = allKeys.filter(k=>sel.has(k)).length;
  if(hint) hint.textContent = totalSel===allKeys.length
    ? `All ${{allKeys.length}} projects shown`
    : `${{totalSel}} of ${{allKeys.length}} selected`;

  // Select All wiring
  const allBox = box.querySelector('#dp-all');
  allBox.indeterminate = !noneChecked && !allChecked;
  allBox.addEventListener('change',()=>{{
    const cur = state.drillSelected ? new Set(state.drillSelected) : new Set(allKeys);
    if(allBox.checked) visKeys.forEach(k=>cur.add(k));
    else visKeys.forEach(k=>cur.delete(k));
    state.drillSelected = cur.size===allKeys.length ? null : cur;
    renderDrillSelector(projects);
    renderDrillDown(projects);
  }});

  // Individual checkbox wiring
  box.querySelectorAll('.dp-cb').forEach(cb=>{{
    cb.addEventListener('change',()=>{{
      const cur = state.drillSelected ? new Set(state.drillSelected) : new Set(allKeys);
      if(cb.checked) cur.add(cb.dataset.key);
      else cur.delete(cb.dataset.key);
      state.drillSelected = cur.size===allKeys.length ? null : cur;
      renderDrillSelector(projects);
      renderDrillDown(projects);
    }});
  }});
}}

// ── Project Drill-Down (column-centric) ────────────────────────────────────
const DRILL_METRICS = METRICS; // all 8 metrics as rows

function renderDrillDown(prebuilt){{
  const dc = state.drillCountry;
  // prebuilt passed from renderDrillSelector to avoid recompute
  let allProjects = prebuilt || (() => {{
    const p = getProjects(dc);
    const sk = state.drillSort;
    p.sort((a,b)=>(b.metrics[sk]||0)-(a.metrics[sk]||0));
    return p;
  }})();

  const sortKey = state.drillSort;
  if(!prebuilt) allProjects.sort((a,b)=>(b.metrics[sortKey]||0)-(a.metrics[sortKey]||0));

  // Apply selection filter
  let projects;
  if(state.drillSelected && state.drillSelected.size > 0){{
    projects = allProjects.filter(p=>state.drillSelected.has(p.name+'||'+p.country));
  }} else {{
    projects = allProjects;
  }}
  const top = projects.slice(0,25);

  document.getElementById('drill-count').textContent =
    state.drillSelected
      ? `${{top.length}} selected of ${{allProjects.length}} projects`
      : `${{top.length}} of ${{allProjects.length}} projects`;

  if(!top.length){{
    document.getElementById('drill-wrap').innerHTML='<div class="no-data">No project data for selected filters.</div>';
    return;
  }}

  // Max per metric for bar widths
  const maxVals = {{}};
  DRILL_METRICS.forEach(m=>{{
    maxVals[m.key] = Math.max(...top.map(p=>p.metrics[m.key]||0),1);
  }});

  let html=`<div class="drill-outer"><table class="drill-table">
    <thead><tr>
      <th class="metric-col">Metric / Project →</th>
      ${{top.map((p,i)=>{{
        const bg=PROJ_BG[i%PROJ_BG.length];
        const bd=PROJ_BORDER[i%PROJ_BORDER.length];
        const cc=COUNTRY_COLORS[p.country];
        return `<th style="background:${{bg}};border-bottom:2px solid ${{bd}}">
          <div class="proj-head">
            <span class="proj-flag" style="background:${{cc}}">${{p.country.substring(0,2).toUpperCase()}}</span>
            <span class="proj-name" title="${{p.name}}">${{p.name}}</span>
          </div>
        </th>`;
      }}).join('')}}
    </tr></thead>
    <tbody>`;

  for(const met of DRILL_METRICS){{
    const isSorted = met.key===sortKey;
    html+=`<tr class="${{isSorted?'sort-row':''}}">
      <td class="metric-col" style="font-weight:${{isSorted?'800':'600'}};color:${{isSorted?met.color:'#334155'}}">
        ${{met.icon}} ${{met.label}}${{isSorted?' ▲':''}}
      </td>`;
    for(let i=0;i<top.length;i++){{
      const p=top[i];
      const v=p.metrics[met.key]||0;
      const barW=v>0?Math.round(v/maxVals[met.key]*100):0;
      const bd=PROJ_BORDER[i%PROJ_BORDER.length];
      const display = v===0?'<span class="cell-num zero">—</span>'
        :`<span class="cell-num">${{met.prefix}}${{Math.round(v).toLocaleString()}}</span>`;
      html+=`<td>
        <div class="cell-inner">
          ${{display}}
          ${{v>0?`<div class="cell-bar-wrap"><div class="cell-bar-fill" style="width:${{barW}}%;background:${{bd}}"></div></div>`:''}}
        </div>
      </td>`;
    }}
    html+=`</tr>`;
  }}

  html+=`</tbody></table></div>`;
  document.getElementById('drill-wrap').innerHTML = html;
}}

// ── Monthly Modal ─────────────────────────────────────────────────────────────
function openModal(yr){{
  const yd=DATA.years[yr];
  if(!yd?.monthly||!Object.keys(yd.monthly).length){{ alert(`Monthly data not available for ${{yr}}`); return; }}
  const months=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const actCtrs = getActiveCountries();
  const singleC = actCtrs.length===1 ? actCtrs[0] : null;
  const getV=(m,k)=>{{
    if(singleC&&yd.countries?.[singleC]?.monthly?.[m]) return yd.countries[singleC].monthly[m][k]||0;
    if(!singleC){{
      // sum across active countries
      return actCtrs.reduce((s,c)=>s+(yd.countries?.[c]?.monthly?.[m]?.[k]||0),0) || (yd.monthly?.[m]?.[k]||0);
    }}
    return yd.monthly?.[m]?.[k]||0;
  }};
  document.getElementById('modal-title').textContent=`Monthly Breakdown — ${{yr}}`;
  document.getElementById('modal-sub').textContent=`${{singleC||'All Countries'}} · Click outside to close`;
  document.getElementById('modal-overlay').classList.add('open');
  const ctx=document.getElementById('modal-chart');
  if(modalChart)modalChart.destroy();
  modalChart=new Chart(ctx,{{
    type:'bar',
    data:{{labels:months,datasets:[
      {{label:'Trained',data:months.map(m=>getV(m,'trained_total')),backgroundColor:'#BFDBFE',borderColor:'#3B82F6',borderWidth:1,borderRadius:3}},
      {{label:'ToT',data:months.map(m=>getV(m,'tot_total')),backgroundColor:'#DDD6FE',borderColor:'#8B5CF6',borderWidth:1,borderRadius:3}},
      {{label:'Biz Supported',data:months.map(m=>getV(m,'biz_supported')),backgroundColor:'#FDE68A',borderColor:'#F59E0B',borderWidth:1,borderRadius:3}},
      {{label:'Jobs',data:months.map(m=>getV(m,'jobs')),backgroundColor:'#FECACA',borderColor:'#EF4444',borderWidth:1,borderRadius:3}},
    ]}},
    options:{{
      responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{position:'top',labels:{{color:'#64748B',font:{{size:11}}}}}},tooltip:{{backgroundColor:'#1E293B',borderColor:'#334155',borderWidth:1}}}},
      scales:{{
        x:{{grid:{{color:'#F1F5F9'}},ticks:{{color:'#64748B'}}}},
        y:{{grid:{{color:'#F1F5F9'}},ticks:{{color:'#64748B'}}}}
      }}
    }}
  }});
}}

// ── Init & Wiring ─────────────────────────────────────────────────────────────
function renderDrill(){{
  const dc = state.drillCountry;
  const allP = getProjects(dc);
  allP.sort((a,b)=>(b.metrics[state.drillSort]||0)-(a.metrics[state.drillSort]||0));
  // reset selection when filters change (year/country/drillCountry)
  state.drillSelected = null;
  state.drillSearch = '';
  const si = document.getElementById('drill-search');
  if(si) si.value='';
  renderDrillSelector(allP);
  renderDrillDown(allP);
}}

function renderAll(){{
  renderKPI();
  updateTrend();
  renderDonut();
  renderCountryTable();
  renderGender();
  renderGrants();
  renderDrill();
}}

// ── Header filter dropdown builder ────────────────────────────────────────────
function buildFilterDropdown({{pillId, panelId, valId, items, stateKey, allLabel, onChange, searchable}}){{
  const pill  = document.getElementById(pillId);
  const panel = document.getElementById(panelId);
  const valEl = document.getElementById(valId);
  if(!pill||!panel||!valEl) return;

  let searchQ = '';

  // items can be an array or a function that returns an array (for dynamic lists)
  const resolveItems = () => typeof items==='function' ? items() : items;

  function getSelected(){{ return state[stateKey]; }}

  function updatePillLabel(){{
    const sel = getSelected();
    const allItems = resolveItems();
    if(!sel||sel.size===0){{
      valEl.textContent = allLabel;
      const badge = pill.querySelector('.fd-badge');
      if(badge) badge.remove();
    }} else {{
      const names = allItems.filter(i=>sel.has(i.key)).map(i=>i.label);
      valEl.textContent = names.length===0 ? allLabel : names.length<=2 ? names.join(', ') : names[0]+' +'+(names.length-1);
      let badge = pill.querySelector('.fd-badge');
      if(!badge){{ badge=document.createElement('span'); badge.className='fd-badge'; pill.insertBefore(badge,panel); }}
      badge.textContent = sel.size;
    }}
  }}

  function buildPanel(){{
    const allItems = resolveItems();
    const visible  = searchQ ? allItems.filter(i=>i.label.toLowerCase().includes(searchQ.toLowerCase())) : allItems;
    const sel = getSelected();
    const allChecked  = !sel || sel.size===allItems.length;
    const someChecked = sel && sel.size>0 && sel.size<allItems.length;

    let h = '';
    if(searchable){{
      h += `<div class="fd-search-wrap"><input type="text" class="fd-search" placeholder="Search…" value="${{searchQ}}" autocomplete="off"></div>`;
    }}
    h += `<label class="fd-row all-row"><input type="checkbox" id="${{pillId}}-all"${{allChecked?' checked':''}}><span>${{allLabel}} (${{allItems.length}})</span></label>`;
    for(const item of visible){{
      const chk = !sel || sel.has(item.key);
      h += `<label class="fd-row"><input type="checkbox" class="fd-cb" data-key="${{item.key}}"${{chk?' checked':''}}><span>${{item.label}}</span></label>`;
    }}
    if(visible.length===0) h += `<div style="padding:10px 14px;font-size:.72rem;color:var(--muted)">No matches</div>`;
    panel.innerHTML = h;

    // Wire search box
    if(searchable){{
      const si = panel.querySelector('.fd-search');
      si.addEventListener('input', e=>{{
        e.stopPropagation();
        searchQ = e.target.value;
        buildPanel();
        const ni = panel.querySelector('.fd-search');
        if(ni){{ ni.focus(); ni.setSelectionRange(ni.value.length,ni.value.length); }}
      }});
      si.addEventListener('click',  e=>e.stopPropagation());
      si.addEventListener('keydown',e=>e.stopPropagation());
    }}

    const allBox = panel.querySelector(`#${{pillId}}-all`);
    allBox.indeterminate = someChecked;
    allBox.addEventListener('change',()=>{{
      state[stateKey] = allBox.checked ? null : new Set();
      searchQ=''; buildPanel(); updatePillLabel(); onChange();
    }});
    panel.querySelectorAll('.fd-cb').forEach(cb=>{{
      cb.addEventListener('change',e=>{{
        e.stopPropagation();
        const cur = getSelected() ? new Set(getSelected()) : new Set(allItems.map(i=>i.key));
        if(cb.checked) cur.add(cb.dataset.key); else cur.delete(cb.dataset.key);
        state[stateKey] = cur.size===allItems.length ? null : cur;
        buildPanel(); updatePillLabel(); onChange();
      }});
    }});
  }}

  // Toggle open/close
  pill.addEventListener('click', e=>{{
    e.stopPropagation();
    const isOpen = panel.classList.contains('open');
    document.querySelectorAll('.fd-panel.open').forEach(p=>{{ p.classList.remove('open'); p.closest('.filter-dd')?.classList.remove('open'); }});
    if(!isOpen){{ searchQ=''; buildPanel(); panel.classList.add('open'); pill.classList.add('open'); }}
  }});
  panel.addEventListener('click',e=>e.stopPropagation());

  updatePillLabel();

  // Return handle for external control
  return {{
    refresh: updatePillLabel,
    pruneInvalid(){{
      const validKeys = new Set(resolveItems().map(i=>i.key));
      const sel = getSelected();
      if(sel){{
        const pruned = new Set([...sel].filter(k=>validKeys.has(k)));
        state[stateKey] = pruned.size===0 ? null : (pruned.size===validKeys.size ? null : pruned);
        updatePillLabel();
      }}
    }}
  }};
}}

// Close dropdowns when clicking outside
document.addEventListener('click',()=>{{
  document.querySelectorAll('.fd-panel.open').forEach(p=>{{ p.classList.remove('open'); p.closest('.filter-dd')?.classList.remove('open'); }});
}});

function init(){{
  document.getElementById('updated-date').textContent=DATA.updated;

  // Project filter (built first so Year can reference it)
  const projFilter = buildFilterDropdown({{
    pillId:'fd-project', panelId:'fdp-project', valId:'fdv-project',
    items: ()=>getProjectNamesForYears(getActiveYears()),  // dynamic: respects active years
    stateKey:'activeProjects', allLabel:'All Projects',
    onChange: renderAll,
    searchable: true
  }});

  // Year filter — when years change, prune invalid project selections then re-render
  buildFilterDropdown({{
    pillId:'fd-year', panelId:'fdp-year', valId:'fdv-year',
    items: getYears().map(y=>( {{key:y, label:y}} )),
    stateKey:'activeYears', allLabel:'All Years',
    onChange: ()=>{{ projFilter.pruneInvalid(); renderAll(); }}
  }});

  // Country filter
  buildFilterDropdown({{
    pillId:'fd-country', panelId:'fdp-country', valId:'fdv-country',
    items: COUNTRIES.map(c=>( {{key:c, label:c}} )),
    stateKey:'activeCountries', allLabel:'All Countries',
    onChange: renderAll
  }});

  // Drill-down controls
  document.getElementById('drill-sort').addEventListener('change',e=>{{ state.drillSort=e.target.value; renderDrill(); }});
  document.getElementById('drill-country').addEventListener('change',e=>{{ state.drillCountry=e.target.value; renderDrill(); }});

  // Drill search
  document.getElementById('drill-search').addEventListener('input',e=>{{
    state.drillSearch=e.target.value;
    const allP = getProjects(state.drillCountry);
    allP.sort((a,b)=>(b.metrics[state.drillSort]||0)-(a.metrics[state.drillSort]||0));
    renderDrillSelector(allP);
  }});

  // Modal
  document.getElementById('modal-close').addEventListener('click',()=>document.getElementById('modal-overlay').classList.remove('open'));
  document.getElementById('modal-overlay').addEventListener('click',e=>{{ if(e.target.id==='modal-overlay') document.getElementById('modal-overlay').classList.remove('open'); }});

  // Initial render
  renderTrendChecks();
  renderCountryChecks();
  initTrend();
  renderKPI();
  renderDonut();
  renderCountryTable();
  renderGender();
  renderGrants();
  renderDrill();
}}

window.addEventListener('DOMContentLoaded', init);
</script>
</body>
</html>"""

# ─── Entry Point ──────────────────────────────────────────────────────────────

def main():
    print("="*50)
    print("  StartHub Africa Impact Dashboard Generator")
    print("="*50)
    if not EXCEL_FILE.exists():
        print(f"\n❌ Excel file not found: {EXCEL_FILE}")
        return
    data = extract_all_data()
    print("\nGenerating HTML dashboard...")
    html = generate_html(data)
    OUTPUT_FILE.write_text(html, encoding='utf-8')
    size_kb = OUTPUT_FILE.stat().st_size // 1024
    print(f"✅ Dashboard saved: {OUTPUT_FILE} ({size_kb} KB)")
    print(f"\nPush index.html to GitHub to publish — or let sync_dashboard.py do it automatically.")
    print("Re-run this script any time you update the Excel file.")

if __name__ == '__main__':
    main()
