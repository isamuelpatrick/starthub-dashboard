# -*- coding: utf-8 -*-
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
    logo_b64 = "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"

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>StartHub Africa — Impact Dashboard</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>

*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
:root{{
  --bg-app:        #f7f8f8;
  --bg-card:       #ffffff;
  --bg-surface:    #f6f8f7;
  --border:        #e5e7eb;
  --border-soft:   #f3f4f6;
  --fg1:           #111827;
  --fg2:           #4b5563;
  --fg3:           #9ca3af;
  --evergreen:     #042a2b;
  --evergreen-mid: #0a4a4d;
  --tangerine:     #ef7b45;
  --tangerine-bg:  #fdf0ea;
  --action:        #0d6e62;
  --action-soft:   #e3f1ee;
  --action-txt:    #0d6e62;
  --radius:        8px;
  --radius-sm:     4px;
  --radius-tag:    6px;
  --radius-card:   12px;
  --radius-pill:   99px;
  --shadow-sm:     0 1px 2px rgba(0,0,0,.05);
  --shadow-card:   0 1px 3px rgba(4,42,43,.06);
  --shadow-hover:  0 4px 12px rgba(4,42,43,.10);
  --shadow-modal:  0 8px 32px rgba(4,42,43,.14);
  --font:          'DM Sans', system-ui, -apple-system, sans-serif;
}}
html{{font-size:14px;color-scheme:light}}
body{{background:var(--bg-app);color:var(--fg1);font-family:var(--font);min-height:100vh;-webkit-font-smoothing:antialiased}}

/* ── Header ── */
#header{{
  background:var(--bg-card);border-bottom:1px solid var(--border);
  padding:6px 24px;min-height:60px;display:flex;align-items:center;gap:16px;
  position:sticky;top:0;z-index:200;box-shadow:var(--shadow-sm);flex-wrap:wrap;
}}
.logo{{display:flex;align-items:center;gap:10px;flex:1;min-width:180px}}
.logo svg{{width:32px;height:32px;flex-shrink:0}}
.logo-text h1{{font-size:14px;font-weight:600;color:var(--fg1);line-height:1.1;letter-spacing:-.01em}}
.logo-text span{{font-size:10px;font-weight:500;color:var(--fg3);letter-spacing:.08em;text-transform:uppercase;display:block;margin-top:2px}}
.header-sep{{width:1px;height:22px;background:var(--border);flex-shrink:0}}
.header-filters{{display:flex;gap:8px;flex-wrap:wrap;align-items:center}}
.filter-pill{{
  display:flex;align-items:center;
  border:1px solid var(--border);border-radius:var(--radius-tag);
  overflow:hidden;background:var(--bg-surface);height:30px;
}}
.filter-pill label{{
  font-size:10px;font-weight:600;color:var(--fg3);
  padding:0 9px;border-right:1px solid var(--border);
  white-space:nowrap;height:100%;display:flex;align-items:center;
  letter-spacing:.07em;text-transform:uppercase;
}}
select{{
  background:transparent;border:none;color:var(--fg1);
  font-size:12px;font-weight:500;font-family:var(--font);
  cursor:pointer;outline:none;padding:0 9px;height:100%;
}}
.updated{{font-size:10px;color:var(--fg3);margin-left:auto;white-space:nowrap;flex-shrink:0}}

/* ── Layout ── */
#main{{padding:22px 24px;max-width:1600px;margin:0 auto}}
.sec-label{{
  font-size:10px;font-weight:600;text-transform:uppercase;
  letter-spacing:.12em;color:var(--evergreen);
  margin-bottom:12px;display:flex;align-items:center;gap:10px;
}}
.sec-label::after{{content:'';flex:1;height:1px;background:var(--border)}}
.mb{{margin-bottom:24px}}

/* ── KPI Cards ── */
#kpi-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(168px,1fr));gap:12px}}
.kpi{{
  background:var(--bg-card);border:1px solid var(--border);border-radius:var(--radius);
  padding:16px 18px 14px;box-shadow:var(--shadow-card);
  border-top:3px solid var(--kpi-accent,var(--evergreen));
  transition:box-shadow .18s,transform .18s;
}}
.kpi:hover{{box-shadow:var(--shadow-hover);transform:translateY(-1px)}}
.kpi-header{{display:flex;align-items:center;gap:7px;margin-bottom:10px}}
.kpi-dot{{width:7px;height:7px;border-radius:50%;background:var(--kpi-accent,var(--evergreen));flex-shrink:0}}
.kpi-label{{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.07em;color:var(--fg3);line-height:1}}
.kpi-value{{font-size:26px;font-weight:600;color:var(--fg1);line-height:1;font-variant-numeric:tabular-nums lining-nums;letter-spacing:-.02em;margin-bottom:8px}}
.kpi-sub{{font-size:10px;color:var(--fg3);display:flex;gap:6px;flex-wrap:wrap;align-items:center}}
.badge{{display:inline-flex;align-items:center;gap:3px;font-size:10px;font-weight:600;padding:2px 8px;border-radius:var(--radius-pill)}}
.badge-m{{background:var(--action-soft);color:var(--action-txt)}}
.badge-f{{background:#f5e3e6;color:#9b3a5a}}

/* ── Chart cards ── */
.chart-card{{background:var(--bg-card);border:1px solid var(--border);border-radius:var(--radius);padding:20px 22px;box-shadow:var(--shadow-card)}}
.chart-card h3{{font-size:14px;font-weight:600;color:var(--fg1);margin-bottom:3px;letter-spacing:-.01em}}
.chart-meta{{font-size:11px;color:var(--fg3);margin-bottom:14px}}
.chart-wrap canvas{{width:100%!important}}
.charts-2{{display:grid;grid-template-columns:2fr 1fr;gap:16px}}
.charts-eq{{display:grid;grid-template-columns:1fr 1fr;gap:16px}}
@media(max-width:900px){{.charts-2,.charts-eq{{grid-template-columns:1fr}}}}

/* ── Metric checkboxes ── */
.mcheck-group{{display:flex;flex-wrap:wrap;gap:4px;padding:8px 10px;background:var(--bg-surface);border:1px solid var(--border);border-radius:var(--radius-tag)}}
.mcheck-item{{display:flex;align-items:center;gap:5px;cursor:pointer;user-select:none;white-space:nowrap;padding:3px 7px;border-radius:4px;transition:background .1s}}
.mcheck-item:hover{{background:var(--bg-card)}}
.mcheck-item input[type=checkbox]{{width:12px;height:12px;cursor:pointer;accent-color:var(--chk-col,var(--evergreen));flex-shrink:0}}
.mcheck-item span{{font-size:11px;font-weight:500;color:var(--fg2)}}
.mcheck-item.all-item span{{font-weight:600;color:var(--fg1)}}
.mcheck-sep{{width:1px;background:var(--border);align-self:stretch;margin:0 2px}}

/* ── Country Metrics Table ── */
.country-metrics-table{{width:100%;border-collapse:collapse;font-size:12px}}
.country-metrics-table th{{padding:7px 12px;text-align:left;font-weight:600;font-size:10px;text-transform:uppercase;letter-spacing:.08em;color:var(--fg3);border-bottom:2px solid var(--border);white-space:nowrap}}
.country-metrics-table td{{padding:9px 12px;border-bottom:1px solid var(--border-soft);white-space:nowrap}}
.country-metrics-table tr:last-child td{{border-bottom:none}}
.country-metrics-table tbody tr:hover td{{background:var(--bg-surface)}}
.c-badge{{display:inline-flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;padding:2px 6px;border-radius:var(--radius-sm);color:#fff;margin-right:6px;letter-spacing:.04em}}
.metric-bar-cell{{display:flex;align-items:center;gap:8px}}
.num-cell{{font-variant-numeric:tabular-nums lining-nums;font-weight:600;color:var(--fg1);text-align:right;min-width:60px;display:inline-block}}
.country-col-head{{display:flex;align-items:center;gap:6px;font-size:12px;font-weight:600}}

/* ── Drill controls ── */
.drill-controls{{display:flex;gap:10px;margin-bottom:14px;align-items:center;flex-wrap:wrap}}
.drill-controls label{{font-size:10px;font-weight:600;color:var(--fg3);text-transform:uppercase;letter-spacing:.06em}}
.ctrl-select{{background:var(--bg-card);border:1px solid var(--border);color:var(--fg1);border-radius:var(--radius-tag);padding:5px 10px;font-size:12px;font-weight:500;font-family:var(--font);cursor:pointer;outline:none;transition:border-color .15s}}
.ctrl-select:focus{{border-color:var(--action)}}
.ctrl-select:hover{{border-color:var(--fg3)}}

/* ── Drill selector ── */
.drill-selector{{display:flex;gap:14px;margin-bottom:14px;align-items:flex-start}}
.drill-selector-panel{{display:flex;flex-direction:column;width:260px;flex-shrink:0;border:1px solid var(--border);border-radius:var(--radius);overflow:hidden;background:var(--bg-card)}}
.drill-search{{width:100%;box-sizing:border-box;padding:8px 10px 8px 30px;border:none;border-bottom:1px solid var(--border);font-size:12px;font-family:var(--font);background:var(--bg-surface);color:var(--fg1);outline:none;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='13' height='13' viewBox='0 0 24 24' fill='none' stroke='%239ca3af' stroke-width='2'%3E%3Ccircle cx='11' cy='11' r='8'/%3E%3Cpath d='m21 21-4.35-4.35'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:9px center}}
.drill-search:focus{{background-color:var(--bg-card)}}
.drill-proj-box{{display:flex;flex-direction:column;overflow-y:auto;max-height:220px;padding:4px 0}}
.dp-item{{display:flex;align-items:center;gap:8px;cursor:pointer;user-select:none;padding:6px 12px;transition:background .1s}}
.dp-item:hover{{background:var(--bg-surface)}}
.dp-item.all-item{{background:var(--bg-surface);border-bottom:1px solid var(--border);padding:7px 12px}}
.dp-item.all-item:hover{{background:var(--action-soft)}}
.dp-item input[type=checkbox]{{width:12px;height:12px;cursor:pointer;accent-color:var(--action);flex-shrink:0}}
.dp-item span{{font-size:11px;font-weight:500;color:var(--fg2);line-height:1.3}}
.dp-item.all-item span{{font-weight:600;color:var(--fg1)}}
.dp-sep{{display:none}}
.dp-none{{font-size:11px;color:var(--fg3);padding:12px}}
.drill-hint{{font-size:10px;color:var(--fg3);padding:5px 12px 7px;border-top:1px solid var(--border);background:var(--bg-surface);text-align:center}}
@media(max-width:700px){{.drill-selector{{flex-direction:column}}.drill-selector-panel{{width:100%}}}}

/* ── Drill table ── */
.drill-outer{{overflow-x:auto;border:1px solid var(--border);border-radius:var(--radius);background:var(--bg-card)}}
.drill-table{{border-collapse:separate;border-spacing:0;table-layout:auto;min-width:100%}}
.drill-table .metric-col{{position:sticky;left:0;z-index:10;background:var(--bg-card);min-width:165px;max-width:165px;border-right:2px solid var(--border);padding:9px 12px;font-size:11px;color:var(--fg2);white-space:nowrap;font-weight:500}}
.drill-table thead .metric-col{{background:var(--bg-surface);font-weight:600;font-size:10px;text-transform:uppercase;letter-spacing:.08em;color:var(--fg3);border-bottom:2px solid var(--border);position:sticky;left:0;top:0;z-index:20}}
.drill-table thead th{{background:var(--bg-surface);border-bottom:2px solid var(--border);padding:8px 12px;min-width:140px;max-width:180px;vertical-align:bottom}}
.proj-head{{display:flex;flex-direction:column;gap:5px}}
.proj-head .proj-name{{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:150px;font-size:11px;font-weight:600;color:var(--fg1)}}
.proj-head .proj-flag{{display:inline-flex;align-items:center;gap:3px;font-size:10px;font-weight:700;padding:2px 6px;border-radius:var(--radius-sm);color:#fff;width:fit-content}}
.drill-table tbody tr:hover .metric-col,.drill-table tbody tr:hover td{{background:#f9fafb}}
.drill-table tbody tr.sort-row .metric-col,.drill-table tbody tr.sort-row td{{background:var(--action-soft)}}
.drill-table tbody td{{padding:8px 12px;border-bottom:1px solid var(--border-soft);vertical-align:middle}}
.cell-inner{{display:flex;flex-direction:column;gap:4px}}
.cell-num{{font-size:12px;font-weight:600;color:var(--fg1);font-variant-numeric:tabular-nums lining-nums;white-space:nowrap}}
.cell-num.zero{{color:var(--fg3);font-weight:400}}
.cell-bar-wrap{{height:3px;background:var(--border);border-radius:2px;overflow:hidden}}
.cell-bar-fill{{height:100%;border-radius:2px;transition:width .3s ease}}

/* ── Misc ── */
.no-data{{text-align:center;color:var(--fg3);padding:32px;font-size:12px}}
#modal-overlay{{display:none;position:fixed;inset:0;background:rgba(4,42,43,.3);z-index:500;align-items:center;justify-content:center}}
#modal-overlay.open{{display:flex}}
#modal{{background:var(--bg-card);border:1px solid var(--border);border-radius:var(--radius-card);padding:28px;width:min(820px,95vw);max-height:90vh;overflow-y:auto;position:relative;box-shadow:var(--shadow-modal)}}
#modal h2{{font-size:15px;font-weight:600;color:var(--fg1);margin-bottom:4px;letter-spacing:-.01em}}
#modal .modal-sub{{font-size:11px;color:var(--fg3);margin-bottom:20px}}
#modal-close{{position:absolute;top:14px;right:16px;background:none;border:none;color:var(--fg3);cursor:pointer;font-size:1rem;width:28px;height:28px;display:flex;align-items:center;justify-content:center;border-radius:var(--radius-sm);transition:background .1s,color .1s}}
#modal-close:hover{{background:var(--bg-surface);color:var(--fg1)}}
::-webkit-scrollbar{{width:4px;height:4px}}
::-webkit-scrollbar-track{{background:transparent}}
::-webkit-scrollbar-thumb{{background:var(--border);border-radius:2px}}
::-webkit-scrollbar-thumb:hover{{background:var(--fg3)}}

/* ── Multi-select ── */
.ms-wrap{{position:relative;display:inline-block}}
.ms-trigger{{display:flex;align-items:center;gap:7px;height:30px;padding:0 10px;background:var(--bg-surface);border:1px solid var(--border);border-radius:var(--radius-tag);cursor:pointer;font-family:var(--font);transition:border-color .15s,background .15s;white-space:nowrap;}}
.ms-trigger:hover{{border-color:var(--fg3);background:var(--bg-card)}}
.ms-wrap.open .ms-trigger{{border-color:var(--action);background:var(--bg-card)}}
.ms-lbl{{font-size:10px;font-weight:600;color:var(--fg3);text-transform:uppercase;letter-spacing:.07em;padding-right:7px;border-right:1px solid var(--border);margin-right:0}}
.ms-val{{font-size:12px;font-weight:500;color:var(--fg1);max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
.ms-badge{{display:inline-flex;align-items:center;justify-content:center;background:var(--action);color:#fff;font-size:9px;font-weight:700;min-width:16px;height:16px;padding:0 4px;border-radius:8px;}}
.ms-chev{{width:10px;height:8px;color:var(--fg3);transition:transform .18s;flex-shrink:0}}
.ms-wrap.open .ms-chev{{transform:rotate(180deg)}}
.ms-dropdown{{display:none;position:absolute;top:calc(100% + 5px);left:0;z-index:400;min-width:210px;max-width:300px;background:var(--bg-card);border:1px solid var(--border);border-radius:var(--radius);box-shadow:var(--shadow-modal);overflow:hidden;}}
.ms-wrap.open .ms-dropdown{{display:block;animation:msDropIn .15s ease}}
@keyframes msDropIn{{from{{opacity:0;transform:translateY(-4px)}}to{{opacity:1;transform:translateY(0)}}}}
.ms-search-wrap{{padding:8px 10px;border-bottom:1px solid var(--border);background:var(--bg-surface)}}
.ms-search{{width:100%;border:1px solid var(--border);border-radius:4px;padding:5px 8px;font-size:12px;font-family:var(--font);color:var(--fg1);background:var(--bg-card);outline:none;box-sizing:border-box}}
.ms-search:focus{{border-color:var(--action)}}
.ms-list{{max-height:240px;overflow-y:auto;padding:4px 0}}
.ms-item{{display:flex;align-items:center;gap:8px;padding:7px 12px;cursor:pointer;user-select:none;transition:background .1s;font-size:12px;color:var(--fg2)}}
.ms-item:hover{{background:var(--bg-surface)}}
.ms-item.all-item{{border-bottom:1px solid var(--border);font-weight:600;color:var(--fg1);background:var(--bg-surface)}}
.ms-item.all-item:hover{{background:var(--action-soft)}}
.ms-item input[type=checkbox]{{width:13px;height:13px;cursor:pointer;accent-color:var(--action);flex-shrink:0}}
.ms-dot{{width:7px;height:7px;border-radius:50%;flex-shrink:0}}
.ms-none{{font-size:11px;color:var(--fg3);padding:10px 12px;text-align:center}}
/* ── Entrance animations ── */
@keyframes fadeUp{{from{{opacity:0;transform:translateY(6px)}}to{{opacity:1;transform:translateY(0)}}}}
@keyframes shimmer{{0%{{background-position:200% 0}}100%{{background-position:-200% 0}}}}
.kpi{{animation:fadeUp .22s ease both}}
.chart-card{{animation:fadeUp .28s ease both}}
.chart-wrap{{position:relative}}
.chart-wrap.loading::before{{content:"";position:absolute;inset:0;border-radius:var(--radius-sm);background:linear-gradient(90deg,var(--border-soft) 0%,#e8ebe9 50%,var(--border-soft) 100%);background-size:200% 100%;animation:shimmer 1.4s ease-in-out infinite;z-index:1;pointer-events:none}}
.chart-wrap.loading canvas{{opacity:0}}
.chart-wrap canvas{{transition:opacity .3s ease}}

</style>
</head>
<body>

<!-- Header -->
<div id="header">
  <div class="logo">
    <img src="data:image/jpeg;base64,{logo_b64}" alt="StartHub Africa" style="height:48px;width:auto;object-fit:contain;flex-shrink:0">
    <div class="logo-text"><h1>StartHub Africa</h1><span>Impact Dashboard</span></div>
  </div>
  <div class="header-sep"></div>
  <div class="header-filters" id="header-filters">
    <div class="ms-wrap" id="ms-year"></div>
    <div class="ms-wrap" id="ms-country"></div>
    <div class="ms-wrap" id="ms-programme"></div>
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
    <div class="chart-meta">Select metrics to compare · Click a year point to see project breakdown</div>
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
        <option value="biz_supported">Businesses Supported</option>
        <option value="biz_started">Businesses Started</option>
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
      <span id="drill-count" style="font-size:.7rem;color:var(--fg3);margin-left:4px"></span>
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
const METRICS = [
  {{key:'trained_total', label:'Trained',              color:'#042a2b', prefix:''}},
  {{key:'tot_total',     label:'Training of Trainers', color:'#3d8f9c', prefix:''}},
  {{key:'biz_supported', label:'Businesses Supported',  color:'#ef7b45', prefix:''}},
  {{key:'biz_started',   label:'Businesses Started',   color:'#1a7a6e', prefix:''}},
  {{key:'jobs',          label:'Jobs Placed',           color:'#f59e0b', prefix:''}},
  {{key:'internships',   label:'Internships',           color:'#5eb1bf', prefix:''}},
  {{key:'grants_count',  label:'Grants',                color:'#d496a7', prefix:''}},
  {{key:'grants_value',  label:'Grant Value',           color:'#754043', prefix:'$'}},
];
const METRIC_MAP = Object.fromEntries(METRICS.map(m=>[m.key,m]));

const COUNTRY_COLORS = {{Uganda:'#042a2b', Tanzania:'#ef7b45', Kenya:'#3d8f9c'}};
const COUNTRY_BG     = {{Uganda:'#e3f3f1', Tanzania:'#fdf0ea', Kenya:'#eaf6f9'}};
const COUNTRIES = ['Uganda','Tanzania','Kenya'];

// Pastel palette for project columns
const PROJ_BG = [
  '#e3f3f1','#fdf0ea','#eaf6f9','#f5e3e6','#fffbeb',
  '#f6f8f7','#e3f3f1','#fdf4f0','#eaf6f9','#f0eef5',
  '#f5f0ea','#eaf6f9','#f0f4f8','#f5e3e6','#e3f5f0',
  '#fdf0ea','#eef5f3','#f5eaef','#e8f0f5','#f8f5ea',
];
/* Solid accent colours for the drill-table mini-bars — one per brand hue */
const PROJ_BORDER = [
  '#1a7a6e','#ef7b45','#3d8f9c','#d496a7','#f59e0b',
  '#042a2b','#5eb1bf','#754043','#0d6e62','#c45e2a',
  '#1a7a6e','#3d8f9c','#ef7b45','#d496a7','#f59e0b',
  '#042a2b','#5eb1bf','#754043','#0d6e62','#c45e2a',
];

// ── State ──────────────────────────────────────────────────────────────────
const state = {{
  years:[], countries:[], programmes:[],
  activeMetrics: METRICS.map(m=>m.key),
  countryMetrics: METRICS.map(m=>m.key),
  drillSort:'trained_total', drillCountry:'all',
  drillSearch:'', drillSelected:null,
}};

let charts = {{}};
let modalChart = null;

// ── Data helpers ───────────────────────────────────────────────────────────
const getYears = () => Object.keys(DATA.years).sort();

function getRawMetrics(yr,country){{
  const d=DATA.years[yr]; if(!d) return null;
  if(country==='all') return d.totals;
  return d.countries?.[country]?.totals||null;
}}

function getMetrics(yr,country){{
  if(!state.programmes.length) return getRawMetrics(yr,country);
  // If all programmes selected, use raw totals for accuracy
  if(state.programmes.length>=getAllProgrammes().length) return getRawMetrics(yr,country);
  // Programme filter: aggregate from project level
  const yd=DATA.years[yr]; if(!yd) return null;
  // Respect both the country argument AND the header country filter
  const headerCtrs = state.countries;
  const ctrs = country==='all' ? headerCtrs : (headerCtrs.includes(country) ? [country] : []);
  const tot=blank();
  for(const c of ctrs){{
    for(const p of (yd.countries?.[c]?.projects||[])){{
      const key=normName(p.name).toLowerCase()+'||'+c;
      if(state.programmes.includes(key)){{
        const m=p.metrics||{{}};
        Object.keys(tot).forEach(k=>{{tot[k]+=(m[k]||0);}});
      }}
    }}
  }}
  return tot;
}}

function getFiltered(){{
  const yrs=state.years;
  const ctrs=state.countries;
  const tot=blank();
  for(const yr of yrs){{
    for(const c of ctrs){{
      const m=getMetrics(yr,c);
      if(m) Object.keys(tot).forEach(k=>{{tot[k]+=(m[k]||0);}});
    }}
  }}
  return tot;
}}

function getAllProgrammes(){{
  const map={{}};
  for(const yr of getYears()){{
    const yd=DATA.years[yr]; if(!yd) continue;
    for(const c of COUNTRIES){{
      for(const p of (yd.countries?.[c]?.projects||[])){{
        const key=normName(p.name).toLowerCase()+'||'+c;
        if(!map[key]) map[key]={{key,name:normName(p.name),country:c}};
      }}
    }}
  }}
  return Object.values(map).sort((a,b)=>a.name.localeCompare(b.name));
}}

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

function getProjects(drillCountry,applyProgFilter){{
  const years = state.years;
  // Resolve effective country list: header country filter + optional drill-country override
  const headerCtrs = state.countries;
  const ctrs = drillCountry !== 'all' ? [drillCountry] : headerCtrs;
  const map = {{}};
  for(const yr of years){{
    const yd = DATA.years[yr]; if(!yd) continue;
    for(const c of ctrs){{
      for(const p of (yd.countries?.[c]?.projects||[])){{
        const key = normName(p.name).toLowerCase()+'||'+c;
        // Apply programme filter when active
        if(state.programmes.length && !state.programmes.includes(key)) continue;
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

// ── CountUp animation ────────────────────────────────────────────────────
const kpiPrev={{}};
function animateKpi(el,target,prefix,key){{
  if(target===0){{el.textContent='\u2014';kpiPrev[key]=0;return;}}
  const prev=kpiPrev[key]!==undefined?kpiPrev[key]:0;
  kpiPrev[key]=target;
  if(prev===target)return;
  const t0=performance.now(),dur=700,sv=prev;
  function step(t){{
    const p=Math.min((t-t0)/dur,1);
    const e=1-Math.pow(1-p,3);
    const cur=Math.round(sv+(target-sv)*e);
    el.textContent=cur<=0?'\u2014':(prefix+cur.toLocaleString());
    if(p<1)requestAnimationFrame(step);
    else el.textContent=target===0?'—':(prefix+target.toLocaleString());
  }}
  requestAnimationFrame(step);
}}

// ── KPI Cards ─────────────────────────────────────────────────────────────
function renderKPI(){{
  const m = getFiltered();
  const rows = METRICS.map((met,idx)=>{{
    const v = m[met.key]||0;
    let sub = '';
    if(met.key==='trained_total'){{
      const tm=m.trained_m||0, tf=m.trained_f||0;
      if(tm||tf) sub=`<span class="badge badge-m">♂ ${{fmtFull(tm)}}</span><span class="badge badge-f">♀ ${{fmtFull(tf)}}</span>`;
    }}
    if(met.key==='tot_total'){{
      const tm=m.tot_m||0, tf=m.tot_f||0;
      if(tm||tf) sub=`<span class="badge badge-m">♂ ${{fmtFull(tm)}}</span><span class="badge badge-f">♀ ${{fmtFull(tf)}}</span>`;
    }}
    if(met.key==='grants_value'){{
      sub=`<span style="color:var(--fg3)">${{m.grants_count||0}} grants awarded</span>`;
    }}
    return '<div class="kpi" style="--kpi-accent:'+met.color+';animation-delay:'+idx*.04+'s">'+
           '<div class="kpi-header"><span class="kpi-dot"></span><span class="kpi-label">'+met.label+'</span></div>'+
           '<div class="kpi-value" data-key="'+met.key+'" data-prefix="'+met.prefix+'">'+fmtFull(v,met.prefix)+'</div>'+
           (sub?'<div class="kpi-sub">'+sub+'</div>':'')+'</div>';
  }});
  document.getElementById('kpi-grid').innerHTML = rows.join('');
  // Animate
  document.querySelectorAll('.kpi-value').forEach(el=>{{
    animateKpi(el,m[el.dataset.key]||0,el.dataset.prefix,el.dataset.key);
  }});
}}

// ── Checkbox group builder (reusable) ────────────────────────────────────────
function buildCheckGroup(containerId, stateKey, onChangeFn){{
  const wrap = document.getElementById(containerId);
  const allChecked = () => state[stateKey].length === METRICS.length;

  wrap.innerHTML = `
    <label class="mcheck-item all-item">
      <input type="checkbox" id="${{containerId}}-all" ${{allChecked()?'checked':''}}>
      <span>Select All</span>
    </label>
    <div class="mcheck-sep"></div>
    ${{METRICS.map(m=>`
      <label class="mcheck-item" style="--chk-col:${{m.color}}">
        <input type="checkbox" data-mk="${{m.key}}" ${{state[stateKey].includes(m.key)?'checked':''}}>
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

// ── Trend Chart ─────────────────────────────────────────────────────────────
function renderTrendChecks(){{
  buildCheckGroup('trend-checks', 'activeMetrics', updateTrend);
}}

function buildTrendDatasets(){{
  // Only show the actively-selected years as chart labels
  const activeYrs = [...state.years].sort();
  const ctrs = state.countries;
  return {{
    labels: activeYrs,
    datasets: METRICS.filter(m=>state.activeMetrics.includes(m.key)).map(m=>{{
      const vals = activeYrs.map(yr=>{{
        let tot=0;
        for(const c of ctrs){{const mm=getMetrics(yr,c);if(mm)tot+=mm[m.key]||0;}}
        return tot||null;
      }});
      return {{
        label:m.label, data:vals,
        borderColor:m.color, backgroundColor:m.color+'22',
        borderWidth:2.5, pointRadius:5, pointHoverRadius:7,
        spanGaps:true, fill:false, tension:0.3,
        yAxisID:m.key==='grants_value'?'y2':'y1'
      }};
    }})
  }};
}}

function initTrend(){{
  const ctx = document.getElementById('trend-chart');
  const wrap = ctx.parentElement;
  wrap.classList.add('loading');
  // Fallback: ensure loading class is removed even if onComplete doesn't fire
  setTimeout(()=>wrap.classList.remove('loading'), 1500);
  charts.trend = new Chart(ctx,{{
    type:'line', data:buildTrendDatasets(),
    options:{{
      responsive:true, maintainAspectRatio:false,
      interaction:{{mode:'index',intersect:false}},
      onClick:(e,els)=>{{ if(els.length && charts.trend) openModal(charts.trend.data.labels[els[0].index]); }},
      animation:{{duration:900,easing:'easeOutQuart',onComplete:()=>wrap.classList.remove('loading')}},
      plugins:{{
        legend:{{display:false}},
        tooltip:{{
          backgroundColor:'#042a2b',borderColor:'#0a4a4d',borderWidth:1,
          titleColor:'#f6f8f7',bodyColor:'#9ca3af',padding:12,cornerRadius:6,
          callbacks:{{label:c=>` ${{c.dataset.label}}: ${{c.parsed.y?.toLocaleString?.()??c.parsed.y}}`}}
        }}
      }},
      scales:{{
        x:{{grid:{{color:'#f3f4f6'}},ticks:{{color:'#9ca3af',font:{{size:11}}}}}},
        y1:{{position:'left',grid:{{color:'#f3f4f6'}},ticks:{{color:'#9ca3af',font:{{size:11}},callback:v=>v>=1000?v/1000+'K':v}}}},
        y2:{{position:'right',grid:{{drawOnChartArea:false}},ticks:{{color:'#9ca3af',font:{{size:11}},callback:v=>'$'+(v>=1000?(v/1000).toFixed(0)+'K':v)}}}}
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
  const yrs=state.years;
  const actCtrs=state.countries;
  // Only include active countries in the chart (no 0-value slices for filtered-out countries)
  const vals=actCtrs.map(c=>yrs.reduce((s,yr)=>{{const m=getMetrics(yr,c);return s+(m?.trained_total||0);}},0));
  const ctx = document.getElementById('donut-chart');
  if(charts.donut) charts.donut.destroy();
  charts.donut = new Chart(ctx,{{
    type:'doughnut',
    data:{{labels:actCtrs,datasets:[{{
      data:vals,
      backgroundColor:actCtrs.map(c=>COUNTRY_COLORS[c]),
      borderColor:'#ffffff', borderWidth:3, hoverOffset:10
    }}]}},
    options:{{
      responsive:true,maintainAspectRatio:false,cutout:'62%',
      interaction:{{mode:'nearest',intersect:true}},
      plugins:{{
        legend:{{position:'bottom',labels:{{color:'#9ca3af',font:{{size:11}},padding:14,usePointStyle:true}}}},
        tooltip:{{
          backgroundColor:'#042a2b',borderColor:'#0a4a4d',borderWidth:1,cornerRadius:6,
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
  const yrs=state.years;
  const actCtrs=state.countries;
  const cData={{}};
  for(const c of actCtrs){{
    const tot=blank();
    for(const yr of yrs){{
      const m=getMetrics(yr,c)||{{}};
      Object.keys(tot).forEach(k=>tot[k]+=(m[k]||0));
    }}
    cData[c]=tot;
  }}

  const visibleMetrics = METRICS.filter(m=>state.countryMetrics.includes(m.key));

  const maxes={{}};
  visibleMetrics.forEach(met=>{{
    maxes[met.key]=Math.max(...actCtrs.map(c=>cData[c][met.key]||0),1);
  }});

  let html='<div style="overflow-x:auto"><table class="country-metrics-table"><thead><tr><th style="min-width:160px">Metric</th>';
  for(const c of actCtrs){{
    html+='<th style="min-width:140px"><div class="country-col-head"><span class="c-badge" style="background:'+COUNTRY_COLORS[c]+'">'+c.substring(0,2).toUpperCase()+'</span>'+c+'</div></th>';
  }}
  html+='</tr></thead><tbody>';

  for(const met of visibleMetrics){{
    html+='<tr><td class="metric-col" style="position:unset;border-right:none;max-width:none">'+met.label+'</td>';
    for(const c of actCtrs){{
      const v=cData[c][met.key]||0;
      const barW=Math.round(v/maxes[met.key]*100);
      html+='<td><div class="metric-bar-cell"><span class="num-cell" style="min-width:60px">'+fmtFull(v,met.prefix)+'</span>'+(v>0?'<div style="flex:1;background:var(--border);border-radius:2px;height:5px"><div style="width:'+barW+'%;height:100%;background:'+COUNTRY_COLORS[c]+';border-radius:2px"></div></div>':'')+'</div></td>';
    }}
    html+='</tr>';
  }}
  html+='</tbody></table></div>';
  document.getElementById('country-table-wrap').innerHTML=html;
}}

function blank(){{
  const b = Object.fromEntries(METRICS.map(m=>[m.key,0]));
  b.trained_m=0; b.trained_f=0; b.tot_m=0; b.tot_f=0;
  return b;
}}

// ── Gender Chart ─────────────────────────────────────────────────────────────
function renderGender(){{
  const years=state.years;
  const ctrs=state.countries;
  const mVals=years.map(yr=>{{let v=0;for(const c of ctrs){{const m=getMetrics(yr,c);if(m)v+=m.trained_m||0;}}return v;}});
  const fVals=years.map(yr=>{{let v=0;for(const c of ctrs){{const m=getMetrics(yr,c);if(m)v+=m.trained_f||0;}}return v;}});
  const ctx = document.getElementById('gender-chart');
  if(charts.gender) charts.gender.destroy();
  charts.gender = new Chart(ctx,{{
    type:'bar',
    data:{{labels:years,datasets:[
      {{label:'Male',  data:mVals,backgroundColor:'#3d8f9c',borderColor:'#3d8f9c',borderWidth:0,borderRadius:4,stack:'g'}},
      {{label:'Female',data:fVals,backgroundColor:'#d496a7',borderColor:'#d496a7',borderWidth:0,borderRadius:4,stack:'g'}},
    ]}},
    options:{{
      responsive:true,maintainAspectRatio:false,
      interaction:{{mode:'index',intersect:false}},
      plugins:{{
        legend:{{position:'top',labels:{{color:'#9ca3af',font:{{size:11}},usePointStyle:true}}}},
        tooltip:{{backgroundColor:'#042a2b',borderColor:'#0a4a4d',borderWidth:1,cornerRadius:6,
          callbacks:{{afterBody:items=>{{const t=items.reduce((s,i)=>s+i.parsed.y,0);return t?[`Total: ${{t.toLocaleString()}}`]:[];}}}}
        }}
      }},
      scales:{{
        x:{{stacked:true,grid:{{color:'#f3f4f6'}},ticks:{{color:'#9ca3af',font:{{size:11}}}}}},
        y:{{stacked:true,grid:{{color:'#f3f4f6'}},ticks:{{color:'#9ca3af',font:{{size:11}},callback:v=>v>=1000?v/1000+'K':v}}}}
      }}
    }}
  }});
}}

// ── Grants Chart ─────────────────────────────────────────────────────────────
function renderGrants(){{
  const ps = getProjects('all',true).filter(p=>(p.metrics.grants_value||0)>0)
    .sort((a,b)=>(b.metrics.grants_value||0)-(a.metrics.grants_value||0)).slice(0,14);
  const wrap = document.getElementById('grants-wrap');
  if(!ps.length){{ wrap.innerHTML='<div class="no-data">No grant data for selected filters</div>'; if(charts.grants){{charts.grants.destroy();charts.grants=null;}} return; }}
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
        backgroundColor:ps.map(p=>COUNTRY_COLORS[p.country]),
        borderColor:ps.map(p=>COUNTRY_COLORS[p.country]),
        borderWidth:0,borderRadius:5
      }}]
    }},
    options:{{
      indexAxis:'y',responsive:true,maintainAspectRatio:false,
      interaction:{{mode:'index',intersect:false}},
      plugins:{{
        legend:{{display:false}},
        tooltip:{{backgroundColor:'#042a2b',borderColor:'#0a4a4d',borderWidth:1,cornerRadius:6,
          callbacks:{{label:c=>`USD ${{Math.round(c.parsed.x).toLocaleString()}} · ${{ps[c.dataIndex].country}}`}}
        }}
      }},
      scales:{{
        x:{{grid:{{color:'#f3f4f6'}},ticks:{{color:'#9ca3af',font:{{size:11}},callback:v=>'$'+(v>=1000?(v/1000).toFixed(0)+'K':v)}}}},
        y:{{grid:{{display:false}},ticks:{{color:'#9ca3af',font:{{size:11}}}}}}
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
       `<span title="${{p.name}}">${{p.name}}<em style="color:var(--fg3);font-style:normal;font-size:.65rem;margin-left:4px">[${{flag}}]</em></span></label>`;
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
      <td class="metric-col" style="font-weight:${{isSorted?'600':'500'}};color:${{isSorted?met.color:'#4b5563'}}">
        ${{met.label}}${{isSorted?' ▲':''}}
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
  const c=state.countries.length===1?state.countries[0]:'all';
  const ctrlabel=state.countries.length===0?'All Countries':state.countries.length===1?state.countries[0]:state.countries.length+' countries';
  const getV=(m,k)=>{{
    if(c!=='all'&&yd.countries?.[c]?.monthly?.[m]) return yd.countries[c].monthly[m][k]||0;
    return yd.monthly?.[m]?.[k]||0;
  }};
  document.getElementById('modal-title').textContent='Monthly Breakdown — '+yr;
  document.getElementById('modal-sub').textContent=ctrlabel+' · Click outside to close';
  document.getElementById('modal-overlay').classList.add('open');
  const ctx=document.getElementById('modal-chart');
  if(modalChart)modalChart.destroy();
  modalChart=new Chart(ctx,{{
    type:'bar',
    data:{{labels:months,datasets:[
      {{label:'Trained',            data:months.map(m=>getV(m,'trained_total')), backgroundColor:'#042a2b',borderColor:'#042a2b',borderWidth:0,borderRadius:4}},
      {{label:'Training of Trainers',data:months.map(m=>getV(m,'tot_total')),     backgroundColor:'#3d8f9c',borderColor:'#3d8f9c',borderWidth:0,borderRadius:4}},
      {{label:'Businesses Supported',data:months.map(m=>getV(m,'biz_supported')), backgroundColor:'#ef7b45',borderColor:'#ef7b45',borderWidth:0,borderRadius:4}},
      {{label:'Jobs Placed',        data:months.map(m=>getV(m,'jobs')),           backgroundColor:'#f59e0b',borderColor:'#f59e0b',borderWidth:0,borderRadius:4}},
    ]}},
    options:{{
      responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{position:'top',labels:{{color:'#9ca3af',font:{{size:11}},usePointStyle:true}}}},tooltip:{{backgroundColor:'#042a2b',borderColor:'#0a4a4d',borderWidth:1,cornerRadius:6}}}},
      scales:{{
        x:{{grid:{{color:'#f3f4f6'}},ticks:{{color:'#9ca3af',font:{{size:11}}}}}},
        y:{{grid:{{color:'#f3f4f6'}},ticks:{{color:'#9ca3af',font:{{size:11}}}}}}
      }}
    }}
  }});
}}

// ── Init & Wiring ─────────────────────────────────────────────────────────────
function renderDrill(){{
  const dc=state.drillCountry;
  const allP=getProjects(dc,false);
  allP.sort((a,b)=>(b.metrics[state.drillSort]||0)-(a.metrics[state.drillSort]||0));
  state.drillSelected=null;
  state.drillSearch='';
  const si=document.getElementById('drill-search');
  if(si) si.value='';
  renderDrillSelector(allP);
  renderDrillDown(allP);
}}

function renderAll(){{
  if(state.years.length===0||state.countries.length===0){{
    document.getElementById('kpi-grid').innerHTML='<div class="no-data" style="grid-column:1/-1;padding:32px;font-size:13px">Select at least one year and country above to view data.</div>';
    document.getElementById('country-table-wrap').innerHTML='';
    document.getElementById('grants-wrap').innerHTML='<canvas id="grants-chart"></canvas>';
    document.getElementById('drill-wrap').innerHTML='';
    document.getElementById('drill-proj-box').innerHTML='';
    document.getElementById('drill-count').textContent='';
    document.getElementById('drill-hint').textContent='';
    if(charts.trend){{charts.trend.data={{labels:[],datasets:[]}};charts.trend.update();}}
    if(charts.donut){{charts.donut.destroy();charts.donut=null;}}
    if(charts.gender){{charts.gender.destroy();charts.gender=null;}}
    if(charts.grants){{charts.grants.destroy();charts.grants=null;}}
    return;
  }}
  renderKPI();
  updateTrend();
  renderDonut();
  renderCountryTable();
  renderGender();
  renderGrants();
  renderDrill();
}}

// ── Multi-select component ───────────────────────────────────────────────
function buildMultiSelect(cfg){{
  const id=cfg.id,label=cfg.label,items=cfg.items,state_arr=cfg.state_arr,onChange=cfg.onChange,searchable=cfg.searchable;
  const wrap=document.getElementById(id);
  if(!wrap)return;
  const isAll=function(){{return state_arr.length===items.length;}};
  const getValLabel=function(){{
    if(state_arr.length===0)return 'None';
    if(isAll())return 'All';
    if(state_arr.length===1){{var it=items.find(function(i){{return i.value===state_arr[0];}});return it?it.label.split('[')[0].trim():state_arr[0];}}
    return state_arr.length+' selected';
  }};
  wrap.innerHTML='';
  var btn=document.createElement('button');
  btn.className='ms-trigger';btn.type='button';
  wrap.appendChild(btn);
  var dd=document.createElement('div');
  dd.className='ms-dropdown';
  if(searchable){{
    var sw=document.createElement('div');sw.className='ms-search-wrap';
    sw.innerHTML='<input class="ms-search" placeholder="Search\u2026" autocomplete="off">';
    dd.appendChild(sw);
  }}
  var list=document.createElement('div');list.className='ms-list';
  dd.appendChild(list);wrap.appendChild(dd);
  function refreshBtn(){{
    var badge=(state_arr.length===0||isAll())?'':('<span class="ms-badge">'+state_arr.length+'</span>');
    btn.innerHTML='<span class="ms-lbl">'+label+'</span><span class="ms-val">'+getValLabel()+'</span>'+badge+'<svg class="ms-chev" viewBox="0 0 10 6" fill="none"><path d="M1 1l4 4 4-4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>';
  }}
  function refreshList(q){{
    q=q||'';
    var vis=q?items.filter(function(i){{return i.label.toLowerCase().includes(q);}}):items;
    var h='<label class="ms-item all-item"><input type="checkbox" class="ms-all-cb"'+(isAll()?' checked':'')+'><span>All '+label+'s</span></label>';
    for(var i=0;i<vis.length;i++){{
      var it=vis[i],chk=state_arr.includes(it.value);
      h+='<label class="ms-item"><input type="checkbox" class="ms-cb" data-val="'+it.value+'"'+(chk?' checked':'')+'>';
      if(it.dot)h+='<span class="ms-dot" style="background:'+it.dot+'"></span>';
      h+='<span>'+it.label+'</span></label>';
    }}
    if(!vis.length)h+='<div class="ms-none">No results</div>';
    list.innerHTML=h;
    var allCb=list.querySelector('.ms-all-cb');
    if(allCb)allCb.indeterminate=state_arr.length>0&&state_arr.length<items.length;
    if(allCb)allCb.addEventListener('change',function(){{
      if(allCb.checked){{items.forEach(function(it){{if(!state_arr.includes(it.value))state_arr.push(it.value);}});}}
      else{{state_arr.length=0;}}
      refreshBtn();refreshList(wrap.querySelector('.ms-search')?wrap.querySelector('.ms-search').value.toLowerCase():'');onChange();
    }});
    list.querySelectorAll('.ms-cb').forEach(function(cb){{
      cb.addEventListener('change',function(e){{
        var v=e.target.dataset.val;
        if(e.target.checked){{if(!state_arr.includes(v))state_arr.push(v);}}
        else{{var ix=state_arr.indexOf(v);if(ix>-1)state_arr.splice(ix,1);}}
        if(allCb){{allCb.checked=state_arr.length===items.length;allCb.indeterminate=state_arr.length>0&&state_arr.length<items.length;}}
        refreshBtn();refreshList(wrap.querySelector('.ms-search')?wrap.querySelector('.ms-search').value.toLowerCase():'');onChange();
      }});
    }});
  }}
  refreshBtn();refreshList();
  // Stop clicks inside the dropdown from bubbling to the document close handler
  dd.addEventListener('click',function(e){{e.stopPropagation();}});
  btn.addEventListener('click',function(e){{
    e.stopPropagation();
    var wasOpen=wrap.classList.contains('open');
    document.querySelectorAll('.ms-wrap.open').forEach(function(w){{w.classList.remove('open');}});
    if(!wasOpen){{
      wrap.classList.add('open');
      if(searchable){{
        var si=wrap.querySelector('.ms-search');
        if(si){{si.value='';refreshList();setTimeout(function(){{si.focus();}},50);}}
      }}
    }}
  }});
}}
document.addEventListener('click',function(){{document.querySelectorAll('.ms-wrap.open').forEach(function(w){{w.classList.remove('open');}});}});

// ── Chart.js defaults (brand) ─────────────────────────────────────────────
Chart.defaults.font.family = "'DM Sans', system-ui, sans-serif";
Chart.defaults.font.size   = 11;
Chart.defaults.color       = '#9ca3af';
// Global interaction: show tooltip on nearest element without requiring exact hit
Chart.defaults.interaction = {{mode:'nearest', intersect:false}};

function init(){{
  document.getElementById('updated-date').textContent=DATA.updated;

  // Year multi-select — pre-select all years so dashboard loads full
  getYears().forEach(function(yr){{if(!state.years.includes(yr))state.years.push(yr);}});
  buildMultiSelect({{id:'ms-year',label:'Year',items:getYears().slice().reverse().map(yr=>({{value:yr,label:yr}})),state_arr:state.years,onChange:function(){{state.drillSelected=null;renderAll();}}}});

  // Country multi-select — pre-select all countries
  COUNTRIES.forEach(function(c){{if(!state.countries.includes(c))state.countries.push(c);}});
  buildMultiSelect({{id:'ms-country',label:'Country',items:COUNTRIES.map(c=>({{value:c,label:c,dot:COUNTRY_COLORS[c]}})),state_arr:state.countries,onChange:function(){{state.drillSelected=null;renderAll();}}}});

  // Programme multi-select — pre-select all programmes
  const progItems=getAllProgrammes().map(p=>({{value:p.key,label:p.name+' ['+p.country.substring(0,2)+']',dot:COUNTRY_COLORS[p.country]}}));
  progItems.forEach(function(p){{if(!state.programmes.includes(p.value))state.programmes.push(p.value);}});
  buildMultiSelect({{id:'ms-programme',label:'Programme',items:progItems,state_arr:state.programmes,onChange:function(){{state.drillSelected=null;renderAll();}},searchable:true}});

  // Drill-down controls
  document.getElementById('drill-sort').addEventListener('change',e=>{{ state.drillSort=e.target.value; renderDrill(); }});
  document.getElementById('drill-country').addEventListener('change',e=>{{ state.drillCountry=e.target.value; renderDrill(); }});

  // Drill search
  document.getElementById('drill-search').addEventListener('input',e=>{{
    state.drillSearch=e.target.value;
    const dc = state.drillCountry;
    const allP = getProjects(dc);
    allP.sort((a,b)=>(b.metrics[state.drillSort]||0)-(a.metrics[state.drillSort]||0));
    renderDrillSelector(allP);
    // don't reset selection or re-render the full table when just typing
  }});

  // Modal
  document.getElementById('modal-close').addEventListener('click',()=>document.getElementById('modal-overlay').classList.remove('open'));
  document.getElementById('modal-overlay').addEventListener('click',e=>{{ if(e.target.id==='modal-overlay') document.getElementById('modal-overlay').classList.remove('open'); }});

  // Charts
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

if(document.readyState==='loading'){{window.addEventListener('DOMContentLoaded',init);}}else{{init();}}
</script>
</body>
</html>"""


# --- Entry Point ---

def main():
    print('='*50)
    print('  StartHub Africa Impact Dashboard Generator')
    print('='*50)
    if not EXCEL_FILE.exists():
        print('Excel file not found: ' + str(EXCEL_FILE))
        return
    data = extract_all_data()
    print('Generating HTML dashboard...')
    html = generate_html(data)
    OUTPUT_FILE.write_text(html, encoding='utf-8')
    size_kb = OUTPUT_FILE.stat().st_size // 1024
    print('Dashboard saved: ' + str(OUTPUT_FILE) + ' (' + str(size_kb) + ' KB)')

if __name__ == '__main__':
    main()
