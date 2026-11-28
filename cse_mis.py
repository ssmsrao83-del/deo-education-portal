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
        "@page{size:A4 landscape;margin:6mm;}"
        "body{font-family:Arial,sans-serif;color:#111;margin:0;padding:0;-webkit-print-color-adjust:exact;print-color-adjust:exact;}"
        ".header{text-align:center;border-bottom:1.5px solid #0f766e;padding-bottom:4px;margin-bottom:6px;}"
        ".header h2{margin:0;font-size:15px;color:#0f766e;}"
        ".header p{margin:2px 0 0;font-size:10px;color:#333;}"
        ".print-btn{background:#0f766e;color:#fff;border:none;padding:6px 12px;font-size:11px;font-weight:bold;border-radius:4px;cursor:pointer;margin-bottom:6px;}"
        "@media print{.print-btn{display:none;}}"
        "h3{font-size:11px;margin:7px 0 3px;color:#1e3a8a;}"
        "table{width:100%;border-collapse:collapse;font-size:9.5px;margin-bottom:7px;}"
        "th,td{border:1px solid #94a3b8;padding:3px 4px;text-align:center;line-height:1.25;}"
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

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Schools", f"{len(leap_df):,}")
    c2.metric("Pending Students", f"{leap_df['Pending Students'].sum():,}")
    c3.metric("Pending Parents", f"{leap_df['Pending Parents'].sum():,}")
    tot_p = leap_df["Target Parents"].sum()
    c4.metric("Parents Achieved %", f"{round((leap_df['Achieved Parents'].sum() / tot_p * 100), 2) if tot_p > 0 else 0}%")

    st.caption("🎨 Color Legend (Parents %): 🔴 < 30% | 🟠 30%–49.9% | 🟡 50%–74.9% | 🟢 ≥ 75% (Sorted Least to Highest)")

    t1, t2, t3 = st.tabs(["🏛️ Division-Wise Abstract", "📍 Mandal-Wise Abstract", "🏫 School-Wise Pending"])

    with t1:
        div_abs = leap_df.groupby("Division", as_index=False).agg(
            Schools=("School Code", "count"),
            Target_Students=("Target Students", "sum"),
            Achieved_Students=("Achieved Students", "sum"),
            Pending_Students=("Pending Students", "sum"),
            Target_Parents=("Target Parents", "sum"),
            Achieved_Parents=("Achieved Parents", "sum"),
            Pending_Parents=("Pending Parents", "sum"),
        )
        div_abs["Students_%"] = (div_abs["Achieved_Students"] / div_abs["Target_Students"].replace(0, np.nan) * 100).round(2).fillna(0)
        div_abs["Parents_%"] = (div_abs["Achieved_Parents"] / div_abs["Target_Parents"].replace(0, np.nan) * 100).round(2).fillna(0)
        div_abs = div_abs.sort_values(by="Parents_%", ascending=True)
        styled_div = div_abs.style.map(color_pct_cell, subset=["Parents_%", "Students_%"]).format({"Students_%": "{:.2f}%", "Parents_%": "{:.2f}%"})
        st.dataframe(styled_div, use_container_width=True, hide_index=True)

    with t2:
        mandal_abs = leap_df.groupby(["Division", "Mandal Name"], as_index=False).agg(
            Schools=("School Code", "count"),
            Target_Students=("Target Students", "sum"),
            Achieved_Students=("Achieved Students", "sum"),
            Pending_Students=("Pending Students", "sum"),
            Target_Parents=("Target Parents", "sum"),
            Achieved_Parents=("Achieved Parents", "sum"),
            Pending_Parents=("Pending Parents", "sum"),
        )
        mandal_abs["Students_%"] = (mandal_abs["Achieved_Students"] / mandal_abs["Target_Students"].replace(0, np.nan) * 100).round(2).fillna(0)
        mandal_abs["Parents_%"] = (mandal_abs["Achieved_Parents"] / mandal_abs["Target_Parents"].replace(0, np.nan) * 100).round(2).fillna(0)
        mandal_abs = mandal_abs.sort_values(by="Parents_%", ascending=True)
        styled_mandal = mandal_abs.style.map(color_pct_cell, subset=["Parents_%", "Students_%"]).format({"Students_%": "{:.2f}%", "Parents_%": "{:.2f}%"})
        st.dataframe(styled_mandal, use_container_width=True, hide_index=True)

    with t3:
        sel_m = st.selectbox("Select Mandal:", [ALL_MNDL] + sorted(leap_df["Mandal Name"].unique().tolist()), key="mis_leap_mndl")
        sch_df = leap_df if sel_m == ALL_MNDL else leap_df[leap_df["Mandal Name"] == sel_m]
        if st.checkbox("Show Only Pending Schools (> 0)", value=True, key="mis_leap_chk"):
            sch_df = sch_df[(sch_df["Pending Students"] > 0) | (sch_df["Pending Parents"] > 0)]
        cols = ["Division", "Mandal Name", "School Code", "School Name", "Management", "Target Students", "Achieved Students", "Pending Students", "Students_%", "Target Parents", "Achieved Parents", "Pending Parents", "Parents_%"]
        sch_sorted = sch_df[cols].sort_values(by=["Parents_%", "Pending Parents"], ascending=[True, False])
        st.dataframe(sch_sorted.style.map(color_pct_cell, subset=["Parents_%", "Students_%"]).format({"Students_%": "{:.2f}%", "Parents_%": "{:.2f}%"}), use_container_width=True, hide_index=True)

    dl1, dl2 = st.columns(2)
    with dl1:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            styled_div.to_excel(writer, sheet_name="Division_Abstract", index=False)
            styled_mandal.to_excel(writer, sheet_name="Mandal_Abstract", index=False)
            sch_sorted.to_excel(writer, sheet_name="School_Wise_Pending", index=False)
        st.download_button("📥 Download Excel Abstract", buf.getvalue(), "LEAP_Abstract.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="mis_leap_xl")

    with dl2:
        today_str = datetime.now().strftime("%d-%m-%Y %I:%M %p")
        d_sch = div_abs["Schools"].sum()
        d_ts, d_as, d_ps = div_abs["Target_Students"].sum(), div_abs["Achieved_Students"].sum(), div_abs["Pending_Students"].sum()
        d_tp, d_ap, d_pp = div_abs["Target_Parents"].sum(), div_abs["Achieved_Parents"].sum(), div_abs["Pending_Parents"].sum()
        d_spct = (d_as / d_ts * 100) if d_ts > 0 else 0
        d_ppct = (d_ap / d_tp * 100) if d_tp > 0 else 0

        div_rows = [f"<tr><td>{i}</td><td style='text-align:left;font-weight:bold;'>{r.Division}</td><td>{r.Schools:,}</td><td>{r.Target_Students:,}</td><td>{r.Achieved_Students:,}</td><td>{r.Pending_Students:,}</td>{print_pct_badge(r._9)}<td>{r.Target_Parents:,}</td><td>{r.Achieved_Parents:,}</td><td>{r.Pending_Parents:,}</td>{print_pct_badge(r._10)}</tr>" for i, r in enumerate(div_abs.itertuples(), 1)]
        div_rows.append(f"<tr class='total-row'><td colspan='2'>DISTRICT TOTAL</td><td>{d_sch:,}</td><td>{d_ts:,}</td><td>{d_as:,}</td><td>{d_ps:,}</td><td>{d_spct:.2f}%</td><td>{d_tp:,}</td><td>{d_ap:,}</td><td>{d_pp:,}</td><td>{d_ppct:.2f}%</td></tr>")

        m_rows = [f"<tr><td>{i}</td><td>{r.Division}</td><td style='text-align:left;font-weight:bold;'>{r._2}</td><td>{r.Schools:,}</td><td>{r.Target_Students:,}</td><td>{r.Achieved_Students:,}</td><td>{r.Pending_Students:,}</td>{print_pct_badge(r._10)}<td>{r.Target_Parents:,}</td><td>{r.Achieved_Parents:,}</td><td>{r.Pending_Parents:,}</td>{print_pct_badge(r._11)}</tr>" for i, r in enumerate(mandal_abs.itertuples(), 1)]
        m_rows.append(f"<tr class='total-row'><td colspan='3'>DISTRICT TOTAL</td><td>{d_sch:,}</td><td>{d_ts:,}</td><td>{d_as:,}</td><td>{d_ps:,}</td><td>{d_spct:.2f}%</td><td>{d_tp:,}</td><td>{d_ap:,}</td><td>{d_pp:,}</td><td>{d_ppct:.2f}%</td></tr>")

        body = (
            "<h3>1. Division-Wise Abstract</h3><table><thead><tr><th>S.No</th><th>Division</th><th>Schools</th><th>Target Stu</th><th>Achieved Stu</th><th>Pending Stu</th><th>Stu %</th><th>Target Par</th><th>Achieved Par</th><th>Pending Par</th><th>Par %</th></tr></thead>"
            f"<tbody>{''.join(div_rows)}</tbody></table>"
            "<h3>2. Mandal-Wise Abstract</h3><table><thead><tr><th>S.No</th><th>Division</th><th>Mandal</th><th>Schools</th><th>Target Stu</th><th>Achieved Stu</th><th>Pending Stu</th><th>Stu %</th><th>Target Par</th><th>Achieved Par</th><th>Pending Par</th><th>Par %</th></tr></thead>"
            f"<tbody>{''.join(m_rows)}</tbody></table>"
        )
        html = build_a4_html("LEAP Downloads Abstract", f"Division: {sel_div} | Mgmt: {sel_mgmt} | {today_str}", body)
        st.download_button("🖨️ Print / Save A4 PDF", html.encode("utf-8"), "LEAP_Print.html", "text/html", use_container_width=True, key="mis_leap_pdf")


def render_exam_report(exam_label, file_path, prefix):
    if not os.path.exists(file_path):
        st.warning(f"⚠️ `{file_path}` repository lo ledu.")
        return

    df_exam = load_exam_marks_data(file_path)
    if df_exam.empty:
        st.error("⚠️ Excel headers check cheyandi.")
        return

    ALL_DIV, ALL_MT, ALL_SM, ALL_MNDL = "All Divisions", "All Management Types", "All School Managements", "All Mandals"
    fc1, fc2, fc3 = st.columns(3)
    with fc1:
        sel_div = st.selectbox("Filter by Division:", [ALL_DIV] + sorted(df_exam["Division"].unique().tolist()), key=f"{prefix}_div")
    with fc2:
        sel_mt = st.selectbox("Filter by Management Type:", [ALL_MT] + sorted(df_exam["Management Type"].unique().tolist()), key=f"{prefix}_mt")
    with fc3:
        sel_sm = st.selectbox("Filter by School Management:", [ALL_SM] + sorted(df_exam["School Management"].unique().tolist()), key=f"{prefix}_sm")

    if sel_div != ALL_DIV:
        df_exam = df_exam[df_exam["Division"] == sel_div]
    if sel_mt != ALL_MT:
        df_exam = df_exam[df_exam["Management Type"] == sel_mt]
    if sel_sm != ALL_SM:
        df_exam = df_exam[df_exam["School Management"] == sel_sm]

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Total Schools", f"{len(df_exam):,}")
    k2.metric("Total Students", f"{df_exam['Total Students'].sum():,}")
    k3.metric("Completed", f"{df_exam['Completed Students'].sum():,}")
    k4.metric("Pending (Incl. Partial)", f"{df_exam['Total Pending Students'].sum():,}")
    tot_s = df_exam["Total Students"].sum()
    k5.metric("Completed %", f"{round((df_exam['Completed Students'].sum() / tot_s * 100), 2) if tot_s > 0 else 0}%")

    st.caption("🎨 Color Legend (Completed %): 🔴 < 30% | 🟠 30%–49.9% | 🟡 50%–74.9% | 🟢 ≥ 75% (Sorted Least to Highest)")

    et1, et2, et3, et4 = st.tabs(["🏛️ Division-Wise", "📍 Mandal-Wise", "🏢 Management-Type Wise", "🏫 School-Wise Pending"])

    with et1:
        div_abs = df_exam.groupby("Division", as_index=False).agg(Schools=("School Code", "count"), Total_Students=("Total Students", "sum"), Completed=("Completed Students", "sum"), Partially_Completed=("Partially Completed Students", "sum"), Not_Started=("Not Started Students", "sum"), Total_Pending=("Total Pending Students", "sum"))
        div_abs["Completed_%"] = (div_abs["Completed"] / div_abs["Total_Students"].replace(0, np.nan) * 100).round(2).fillna(0)
        div_abs = div_abs.sort_values(by="Completed_%", ascending=True)
        st.dataframe(div_abs.style.map(color_pct_cell, subset=["Completed_%"]).format({"Completed_%": "{:.2f}%"}), use_container_width=True, hide_index=True)

    with et2:
        mandal_abs = df_exam.groupby(["Division", "Mandal Name"], as_index=False).agg(Schools=("School Code", "count"), Total_Students=("Total Students", "sum"), Completed=("Completed Students", "sum"), Partially_Completed=("Partially Completed Students", "sum"), Not_Started=("Not Started Students", "sum"), Total_Pending=("Total Pending Students", "sum"))
        mandal_abs["Completed_%"] = (mandal_abs["Completed"] / mandal_abs["Total_Students"].replace(0, np.nan) * 100).round(2).fillna(0)
        mandal_abs = mandal_abs.sort_values(by="Completed_%", ascending=True)
        st.dataframe(mandal_abs.style.map(color_pct_cell, subset=["Completed_%"]).format({"Completed_%": "{:.2f}%"}), use_container_width=True, hide_index=True)

    with et3:
        mgmt_abs = df_exam.groupby(["Management Type", "School Management"], as_index=False).agg(Schools=("School Code", "count"), Total_Students=("Total Students", "sum"), Completed=("Completed Students", "sum"), Partially_Completed=("Partially Completed Students", "sum"), Not_Started=("Not Started Students", "sum"), Total_Pending=("Total Pending Students", "sum"))
        mgmt_abs["Completed_%"] = (mgmt_abs["Completed"] / mgmt_abs["Total_Students"].replace(0, np.nan) * 100).round(2).fillna(0)
        mgmt_abs = mgmt_abs.sort_values(by="Completed_%", ascending=True)
        st.dataframe(mgmt_abs.style.map(color_pct_cell, subset=["Completed_%"]).format({"Completed_%": "{:.2f}%"}), use_container_width=True, hide_index=True)

    with et4:
        sel_m = st.selectbox("Select Mandal:", [ALL_MNDL] + sorted(df_exam["Mandal Name"].unique().tolist()), key=f"{prefix}_mndl")
        sch_e = df_exam if sel_m == ALL_MNDL else df_exam[df_exam["Mandal Name"] == sel_m]
        if st.checkbox("Show Only Pending Schools (> 0)", value=True, key=f"{prefix}_chk"):
            sch_e = sch_e[sch_e["Total Pending Students"] > 0]
        cols = ["Division", "Mandal Name", "School Code", "School Name", "Management Type", "School Management", "Total Students", "Completed Students", "Partially Completed Students", "Not Started Students", "Total Pending Students", "Completed_%"]
        sch_sorted = sch_e[cols].sort_values(by=["Completed_%", "Total Pending Students"], ascending=[True, False])
        st.dataframe(sch_sorted.style.map(color_pct_cell, subset=["Completed_%"]).format({"Completed_%": "{:.2f}%"}), use_container_width=True, hide_index=True)

    sdl1, sdl2 = st.columns(2)
    with sdl1:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            div_abs.to_excel(writer, sheet_name="Division_Abstract", index=False)
            mandal_abs.to_excel(writer, sheet_name="Mandal_Abstract", index=False)
            mgmt_abs.to_excel(writer, sheet_name="Management_Type_Abstract", index=False)
            sch_sorted.to_excel(writer, sheet_name="School_Wise_Pending", index=False)
        st.download_button(f"📥 Download {exam_label} Excel (4 Sheets)", buf.getvalue(), f"{exam_label}_Pending.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key=f"{prefix}_xl")

    with sdl2:
        today_str = datetime.now().strftime("%d-%m-%Y %I:%M %p")
        t_sch, t_stu, t_cmp, t_prt, t_nst, t_pnd = div_abs["Schools"].sum(), div_abs["Total_Students"].sum(), div_abs["Completed"].sum(), div_abs["Partially_Completed"].sum(), div_abs["Not_Started"].sum(), div_abs["Total_Pending"].sum()
        t_cpct = (t_cmp / t_stu * 100) if t_stu > 0 else 0

        d_rows = [f"<tr><td>{i}</td><td style='text-align:left;font-weight:bold;'>{r.Division}</td><td>{r.Schools:,}</td><td>{r.Total_Students:,}</td><td>{r.Completed:,}</td><td>{r.Partially_Completed:,}</td><td>{r.Not_Started:,}</td><td>{r.Total_Pending:,}</td>{print_pct_badge(r._8)}</tr>" for i, r in enumerate(div_abs.itertuples(), 1)]
        d_rows.append(f"<tr class='total-row'><td colspan='2'>DISTRICT TOTAL</td><td>{t_sch:,}</td><td>{t_stu:,}</td><td>{t_cmp:,}</td><td>{t_prt:,}</td><td>{t_nst:,}</td><td>{t_pnd:,}</td><td>{t_cpct:.2f}%</td></tr>")

        mg_rows = [f"<tr><td>{i}</td><td style='text-align:left;'>{r._1}</td><td style='text-align:left;font-weight:bold;'>{r._2}</td><td>{r.Schools:,}</td><td>{r.Total_Students:,}</td><td>{r.Completed:,}</td><td>{r.Partially_Completed:,}</td><td>{r.Not_Started:,}</td><td>{r.Total_Pending:,}</td>{print_pct_badge(r._9)}</tr>" for i, r in enumerate(mgmt_abs.itertuples(), 1)]
        mg_rows.append(f"<tr class='total-row'><td colspan='3'>DISTRICT TOTAL</td><td>{t_sch:,}</td><td>{t_stu:,}</td><td>{t_cmp:,}</td><td>{t_prt:,}</td><td>{t_nst:,}</td><td>{t_pnd:,}</td><td>{t_cpct:.2f}%</td></tr>")

        mn_rows = [f"<tr><td>{i}</td><td>{r.Division}</td><td style='text-align:left;font-weight:bold;'>{r._2}</td><td>{r.Schools:,}</td><td>{r.Total_Students:,}</td><td>{r.Completed:,}</td><td>{r.Partially_Completed:,}</td><td>{r.Not_Started:,}</td><td>{r.Total_Pending:,}</td>{print_pct_badge(r._9)}</tr>" for i, r in enumerate(mandal_abs.itertuples(), 1)]
        mn_rows.append(f"<tr class='total-row'><td colspan='3'>DISTRICT TOTAL</td><td>{t_sch:,}</td><td>{t_stu:,}</td><td>{t_cmp:,}</td><td>{t_prt:,}</td><td>{t_nst:,}</td><td>{t_pnd:,}</td><td>{t_cpct:.2f}%</td></tr>")

        body = (
            "<h3>1. Division-Wise Abstract</h3><table><thead><tr><th>S.No</th><th>Division</th><th>Schools</th><th>Total</th><th>Completed</th><th>Partial</th><th>Not Started</th><th>Pending</th><th>Comp %</th></tr></thead>"
            f"<tbody>{''.join(d_rows)}</tbody></table>"
            "<h3>2. Management-Type Wise Abstract</h3><table><thead><tr><th>S.No</th><th>Mgmt Type</th><th>School Mgmt</th><th>Schools</th><th>Total</th><th>Completed</th><th>Partial</th><th>Not Started</th><th>Pending</th><th>Comp %</th></tr></thead>"
            f"<tbody>{''.join(mg_rows)}</tbody></table>"
            "<h3>3. Mandal-Wise Abstract</h3><table><thead><tr><th>S.No</th><th>Division</th><th>Mandal</th><th>Schools</th><th>Total</th><th>Completed</th><th>Partial</th><th>Not Started</th><th>Pending</th><th>Comp %</th></tr></thead>"
            f"<tbody>{''.join(mn_rows)}</tbody></table>"
        )
        html = build_a4_html(f"{exam_label} Marks Entry Pending Abstract", f"Division: {sel_div} | Mgmt: {sel_mt} | {today_str}", body)
        st.download_button(f"🖨️ Print / Save {exam_label} A4 PDF", html.encode("utf-8"), f"{exam_label}_Print.html", "text/html", use_container_width=True, key=f"{prefix}_pdf")

@st.fragment
def render_cse_mis_tab():
    st.subheader("📄 CSE MIS Reports")
    report_choice = st.radio(
        "Select Report:",
        ["📱 LEAP Downloads Abstract", "📝 SA-1 Marks Entry Pending", "📋 FA-2 Marks Entry Pending"],
        horizontal=True,
        key="cse_mis_lazy_selector"
    )
    st.markdown("---")
    if report_choice == "📱 LEAP Downloads Abstract":
        render_leap_report()
    elif report_choice == "📝 SA-1 Marks Entry Pending":
        render_exam_report("SA-1", "SA1- MARKS UPDATION.xlsx", "mis_sa1")
    elif report_choice == "📋 FA-2 Marks Entry Pending":
        render_exam_report("FA-2", "FA2- MARKS UPDATION.xlsx", "mis_fa2")
