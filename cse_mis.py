import os
import io
from datetime import datetime
import numpy as np
import pandas as pd
import streamlit as st

DIVISION_MAP = {
    # BHIMAVARAM DIVISION (10 Mandals)
    "AKIVEEDU": "BHIMAVARAM",
    "ATTILI": "BHIMAVARAM",
    "BHIMAVARAM": "BHIMAVARAM",
    "BHIMAVARAM URBAN": "BHIMAVARAM",
    "BHIMAVARAM RURAL": "BHIMAVARAM",
    "GANAPAVARAM": "BHIMAVARAM",
    "KALLA": "BHIMAVARAM",
    "PALAKODERU": "BHIMAVARAM",
    "PENTAPADU": "BHIMAVARAM",
    "TADEPALLI GUDEM": "BHIMAVARAM",
    "TADEPALLIGUDEM": "BHIMAVARAM",
    "UNDI": "BHIMAVARAM",
    "VEERAVASARAM": "BHIMAVARAM",

    # NARASAPURAM DIVISION (10 Mandals)
    "ACHANTA": "NARASAPURAM",
    "ELAMANCHILI": "NARASAPURAM",
    "YELAMANCHILI": "NARASAPURAM",
    "IRAGAVARAM": "NARASAPURAM",
    "MOGALTUR": "NARASAPURAM",
    "MOGALTHUR": "NARASAPURAM",
    "MOGALTHURU": "NARASAPURAM",
    "NARSAPURAM": "NARASAPURAM",
    "NARASAPURAM": "NARASAPURAM",
    "NARSAPUR": "NARASAPURAM",
    "PALACOLE": "NARASAPURAM",
    "PALAKOLLU": "NARASAPURAM",
    "PENUGONDA": "NARASAPURAM",
    "PENUMANTRA": "NARASAPURAM",
    "PODURU": "NARASAPURAM",
    "TANUKU": "NARASAPURAM",
}


def get_division(mandal_name):
    if pd.isna(mandal_name):
        return "Unmapped"
    clean = " ".join(str(mandal_name).upper().replace("(M)", "").replace("(U)", "").replace("(R)", "").split())
    return DIVISION_MAP.get(clean, "Unmapped")


def color_pct_cell(val):
    try:
        v = float(val)
    except (ValueError, TypeError):
        return ""
    if v < 30:
        return "background-color: #7f1d1d; color: #fecaca; font-weight: bold;"
    elif v < 50:
        return "background-color: #9a3412; color: #fed7aa; font-weight: bold;"
    elif v < 75:
        return "background-color: #854d0e; color: #fef08a; font-weight: bold;"
    else:
        return "background-color: #14532d; color: #bbf7d0; font-weight: bold;"


def print_pct_badge(v):
    if v < 30:
        return f'<td style="background:#fee2e2; color:#991b1b; font-weight:bold;">{v:.2f}%</td>'
    elif v < 50:
        return f'<td style="background:#ffedd5; color:#9a3412; font-weight:bold;">{v:.2f}%</td>'
    elif v < 75:
        return f'<td style="background:#fef9c3; color:#854d0e; font-weight:bold;">{v:.2f}%</td>'
    else:
        return f'<td style="background:#dcfce7; color:#166534; font-weight:bold;">{v:.2f}%</td>'


def build_a4_html(report_title, filter_meta, body_html):
    css = (
        "@page{size:A4 landscape;margin:5mm;}"
        "body{font-family:Arial,sans-serif;color:#111;margin:0;padding:4px;-webkit-print-color-adjust:exact;print-color-adjust:exact;}"
        ".header{text-align:center;border-bottom:1.5px solid #0f766e;padding-bottom:3px;margin-bottom:4px;}"
        ".header h2{margin:0;font-size:13px;color:#0f766e;}"
        ".header p{margin:1px 0 0;font-size:9px;color:#444;}"
        ".print-btn{background:#0f766e;color:#fff;border:none;padding:6px 12px;font-size:11px;font-weight:bold;border-radius:4px;cursor:pointer;margin-bottom:6px;}"
        "@media print{.print-btn{display:none;}body{zoom:0.78;}}"
        "h3{font-size:10px;margin:4px 0 2px;color:#1e3a8a;}"
        "table{width:100%;border-collapse:collapse;font-size:8.2px;margin-bottom:5px;}"
        "th,td{border:1px solid #94a3b8;padding:1.8px 3px;text-align:center;line-height:1.15;}"
        "th{background-color:#0f766e;color:#fff;font-weight:bold;}"
        ".total-row td{background-color:#e2e8f0;font-weight:bold;}"
    )
    return (
        f"<!DOCTYPE html><html><head><meta charset='utf-8'><title>{report_title}</title>"
        f"<style>{css}</style></head><body onload='window.print()'>"
        "<button class='print-btn' onclick='window.print()'>🖨️ Print / Save as PDF</button>"
        "<div class='header'><h2>OFFICE OF THE DISTRICT EDUCATIONAL OFFICER, WEST GODAVARI, BHIMAVARAM</h2>"
        f"<p><b>{report_title}</b> | {filter_meta}</p></div>"
        f"{body_html}</body></html>"
    )


@st.cache_data(ttl=300)
def load_leap_data(file_path):
    preview = pd.read_excel(file_path, header=None, nrows=5)
    header_row = 0
    for idx, row in preview.iterrows():
        row_str = " ".join([str(x).upper() for x in row.values if pd.notna(x)])
        if "SCHOOL CODE" in row_str or "MANDAL NAME" in row_str:
            header_row = idx
            break

    df_raw = pd.read_excel(file_path, header=header_row)
    df_raw.columns = [str(c).strip() for c in df_raw.columns]

    col_map = {c.upper(): c for c in df_raw.columns}
    sch_code_col = col_map.get("SCHOOL CODE")
    mandal_col = col_map.get("MANDAL NAME")
    sch_name_col = col_map.get("SCHOOL NAME")
    mgmt_col = col_map.get("MANAGEMENT")
    div_col = col_map.get("DIVISION")

    if not sch_code_col or not mandal_col:
        return pd.DataFrame()

    df_raw = df_raw[df_raw[sch_code_col].notna()].copy()
    df_raw["School Code"] = df_raw[sch_code_col].astype(str).str.replace(".0", "", regex=False).str.strip()
    df_raw["Mandal Name"] = df_raw[mandal_col].astype(str).str.strip()
    df_raw["School Name"] = df_raw[sch_name_col].astype(str).str.strip() if sch_name_col else ""
    df_raw["Management"] = df_raw[mgmt_col].astype(str).str.strip() if mgmt_col else "All"

    if div_col:
        df_raw["Division"] = df_raw[div_col].fillna("").astype(str).str.strip().str.upper()
        df_raw.loc[df_raw["Division"] == "", "Division"] = df_raw.loc[df_raw["Division"] == "", "Mandal Name"].apply(get_division)
    else:
        df_raw["Division"] = df_raw["Mandal Name"].apply(get_division)

    for target_c in ["Target Students", "Target Parents", "Achieved Students", "Achieved Parents"]:
        actual_c = col_map.get(target_c.upper())
        df_raw[target_c] = pd.to_numeric(df_raw[actual_c], errors="coerce").fillna(0).astype(int) if actual_c else 0

    df_raw["Pending Students"] = (df_raw["Target Students"] - df_raw["Achieved Students"]).clip(lower=0)
    df_raw["Pending Parents"] = (df_raw["Target Parents"] - df_raw["Achieved Parents"]).clip(lower=0)
    df_raw["Students_%"] = (df_raw["Achieved Students"] / df_raw["Target Students"].replace(0, np.nan) * 100).round(2).fillna(0)
    df_raw["Parents_%"] = (df_raw["Achieved Parents"] / df_raw["Target Parents"].replace(0, np.nan) * 100).round(2).fillna(0)
    return df_raw


@st.cache_data(ttl=300)
def load_exam_marks_data(file_path):
    preview = pd.read_excel(file_path, header=None, nrows=5)
    header_row = 0
    for idx, row in preview.iterrows():
        row_str = " ".join([str(x).upper() for x in row.values if pd.notna(x)])
        if "SCHOOL CODE" in row_str or "TOTAL STUDENTS" in row_str:
            header_row = idx
            break

    df_raw = pd.read_excel(file_path, header=header_row)
    df_raw.columns = [str(c).strip() for c in df_raw.columns]

    col_map = {c.upper(): c for c in df_raw.columns}
    sch_code_col = col_map.get("SCHOOL CODE")
    mandal_col = col_map.get("MANDAL NAME")
    sch_name_col = col_map.get("SCHOOL NAME")
    sch_mgmt_col = col_map.get("SCHOOL MANAGEMENT")
    mgmt_type_col = col_map.get("MANAGEMENT TYPE")
    div_col = col_map.get("DIVISION")

    if not sch_code_col or not mandal_col:
        return pd.DataFrame()

    df_raw = df_raw[df_raw[sch_code_col].notna()].copy()
    df_raw["School Code"] = df_raw[sch_code_col].astype(str).str.replace(".0", "", regex=False).str.strip()
    df_raw["Mandal Name"] = df_raw[mandal_col].astype(str).str.strip()
    df_raw["School Name"] = df_raw[sch_name_col].astype(str).str.strip() if sch_name_col else ""
    df_raw["School Management"] = df_raw[sch_mgmt_col].astype(str).str.strip() if sch_mgmt_col else "Unknown"
    df_raw["Management Type"] = df_raw[mgmt_type_col].astype(str).str.strip() if mgmt_type_col else "Unknown"

    if div_col:
        df_raw["Division"] = df_raw[div_col].fillna("").astype(str).str.strip().str.upper()
        df_raw.loc[df_raw["Division"] == "", "Division"] = df_raw.loc[df_raw["Division"] == "", "Mandal Name"].apply(get_division)
    else:
        df_raw["Division"] = df_raw["Mandal Name"].apply(get_division)

    for target_c in ["Total Students", "Completed Students", "Partially Completed Students"]:
        actual_c = col_map.get(target_c.upper())
        df_raw[target_c] = pd.to_numeric(df_raw[actual_c], errors="coerce").fillna(0).astype(int) if actual_c else 0

    df_raw["Not Started Students"] = (df_raw["Total Students"] - df_raw["Completed Students"] - df_raw["Partially Completed Students"]).clip(lower=0)
    df_raw["Total Pending Students"] = (df_raw["Total Students"] - df_raw["Completed Students"]).clip(lower=0)
    df_raw["Completed_%"] = (df_raw["Completed Students"] / df_raw["Total Students"].replace(0, np.nan) * 100).round(2).fillna(0)
    return df_raw


def render_leap_report():
    file_path = "LEAP DOWNLOADS ABSTRACT.xlsx"
    if not os.path.exists(file_path):
        st.warning(f"⚠️ `{file_path}` repository lo ledu.")
        return

    leap_df = load_leap_data(file_path)
    if leap_df.empty:
        st.error("⚠️ Excel headers check cheyandi.")
        return

    ALL_DIV, ALL_MGMT, ALL_MNDL = "All Divisions", "All Managements", "All Mandals"
    f1, f2 = st.columns(2)
    with f1:
        sel_div = st.selectbox("Filter by Division:", [ALL_DIV] + sorted(leap_df["Division"].unique().tolist()), key="mis_leap_div")
    with f2:
        sel_mgmt = st.selectbox("Filter by Management:", [ALL_MGMT] + sorted(leap_df["Management"].unique().tolist()), key="mis_leap_mgmt")

    if sel_div != ALL_DIV:
        leap_df = leap_df[leap_df["Division"] == sel_div]
    if sel_mgmt != ALL_MGMT:
        leap_df = leap_df[leap_df["Management"] == sel_mgmt]

    c1, c2,
