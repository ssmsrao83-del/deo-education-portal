import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import os
import io

st.set_page_config(page_title="District Education Portal", layout="wide")

st.title("🎓 District Educational Office - Management Portal")
st.caption("West Godavari District - School Education Department")
st.markdown("---")

EXCEL_FILE_PATH = "UPTO DATE UDISE ROLL.xlsx"
TEACHERS_FILE_PATH = "TEACHERS DATA.xlsx"
MBU_FILE_PATH = "MBU SCHOOL WISE PENDING.xlsx"

MANAGEMENT_MAPPING = {
    10: "10 - State Govt.",
    24: "24 - APSWREI Society Schools",
    33: "33 - MPP_ZPP SCHOOLS",
    34: "34 - MUNCIPAL",
    35: "35 - Pvt.Aided",
    37: "37 - Pvt.Aided Oriental Schools",
    38: "38 - Pvt.Unaided",
    39: "39 - Pvt.Unaided (CBSE Syllabus)",
    40: "40 - Pvt.Unaided (ICSE Syllabus)",
    42: "42 - Pvt.Unaided (Deaf and Dumb)",
    44: "44 - Pvt.Unaided (Mentally Retarded)",
    66: "66 - BC Welfare",
    67: "67 - Private Unaided(Opening Permission)"
}

CATEGORY_MAPPING = {
    1: "1 - Primary",
    2: "2 - Primary with Upper Primary",
    3: "3 - Pr. with Up.Pr. sec. and H.Sec.",
    5: "5 - Up. Pr. Secondary and Higher Sec",
    6: "6 - Pr. Up Pr. and Secondary Only",
    7: "7 - Upper Pr. and Secondary",
    11: "11 - Higher Secondary only/Jr. College"
}

def get_merged_mgmt(val):
    try:
        val_str = str(val).split('-')[0].strip()
        c = int(float(val_str))
    except:
        c = 0
    if c in [10, 24, 33, 34, 66]:
        return "STATE GOVT"
    elif c in [35, 37]:
        return "AIDED"
    elif c in [38, 39, 40, 42, 44, 67]:
        return "PRIVATE"
    return "OTHER"

def clean_and_map(val, mapping_dict):
    try:
        val_str = str(val).split('-')[0].strip()
        num = int(float(val_str))
        return mapping_dict.get(num, str(val))
    except:
        return str(val)

def append_total_row(df_in, label_col, total_label="TOTAL"):
    if df_in.empty:
        return df_in
    df_calc = df_in.copy()
    total_dict = {}
    for col in df_calc.columns:
        if col == label_col:
            total_dict[col] = total_label
        elif pd.api.types.is_numeric_dtype(df_calc[col]):
            total_dict[col] = df_calc[col].sum()
        else:
            try:
                num_s = pd.to_numeric(df_calc[col], errors='coerce')
                if num_s.notnull().any():
                    total_dict[col] = int(num_s.fillna(0).sum())
                else:
                    total_dict[col] = ""
            except:
                total_dict[col] = ""
    return pd.concat([df_calc, pd.DataFrame([total_dict])], ignore_index=True)

def render_print_button(dataframe, report_title="DEO REPORT", subtitle="West Godavari District"):
    html_table = dataframe.to_html(index=False, classes='print-table')
    html_table_escaped = html_table.replace("`", "'").replace("\\", "\\\\")

    html_code = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <style>
      .p-btn {{
          background-color: #0d6efd;
          color: white;
          border: none;
          padding: 8px 18px;
          border-radius: 6px;
          font-size: 14px;
          font-weight: 600;
          cursor: pointer;
          display: inline-flex;
          align-items: center;
          gap: 6px;
          box-shadow: 0 2px 4px rgba(0,0,0,0.15);
      }}
      .p-btn:hover {{ background-color: #0b5ed7; }}
    </style>
    </head>
    <body style="margin: 0; padding: 0;">
      <button class="p-btn" onclick="printReport()">🖨️ Direct Print / Save as PDF</button>
      <script>
      function printReport() {{
          var tableHtml = `{html_table_escaped}`;
          var printWin = window.open('', '', 'width=1150,height=750');
          printWin.document.write(`
            <html>
            <head>
              <title>{report_title}</title>
              <style>
                body {{ font-family: Arial, sans-serif; margin: 15px; color: #000; }}
                .header-area {{ text-align: center; border-bottom: 2px solid #000; padding-bottom: 6px; margin-bottom: 12px; }}
                .header-area h2 {{ margin: 0; font-size: 16px; color: #1a237e; text-transform: uppercase; }}
                .header-area h3 {{ margin: 3px 0; font-size: 13px; color: #333; }}
                .header-area p {{ margin: 2px 0 0 0; font-size: 11px; color: #555; font-weight: bold; }}
                .print-table {{ width: 100%; border-collapse: collapse; font-size: 10px; }}
                .print-table th, .print-table td {{ border: 1px solid #444; padding: 4px 5px; text-align: center; }}
                .print-table th {{ background-color: #f2f2f2 !important; font-weight: bold; -webkit-print-color-adjust: exact; }}
                .print-table tr:nth-child(even) {{ background-color: #fafafa; -webkit-print-color-adjust: exact; }}
                .print-table tr:last-child {{ font-weight: bold; background-color: #eaeaea !important; border-top: 2px solid #222; }}
                @media print {{ body {{ margin: 8mm; }} }}
              </style>
            </head>
            <body>
              <div class="header-area">
                <h2>GOVERNMENT OF ANDHRA PRADESH - SCHOOL EDUCATION DEPARTMENT</h2>
                <h3>DISTRICT EDUCATIONAL OFFICE - WEST GODAVARI</h3>
                <p>{report_title} | {subtitle}</p>
              </div>
              ${{tableHtml}}
            </body>
            </html>
          `);
          printWin.document.close();
          printWin.focus();
          setTimeout(function() {{
              printWin.print();
              printWin.close();
          }}, 400);
      }}
      </script>
    </body>
    </html>
    """
    components.html(html_code, height=45, scrolling=False)

@st.cache_data(ttl=30)
def load_udise_data(file_path):
    actual_path = None
    for f in os.listdir('.'):
        if f.lower() == file_path.lower():
            actual_path = f
            break
    if not actual_path:
        return None
    try:
        df = pd.read_excel(actual_path)
        df.columns = [str(c).strip() for c in df.columns]
        for col in df.columns:
            if 'MANAGE' in col.upper():
                df['Management_Display'] = df[col].apply(lambda x: clean_and_map(x, MANAGEMENT_MAPPING))
                df['Merged_Management'] = df[col].apply(get_merged_mgmt)
            if 'CATEG' in col.upper():
                df['Category_Display'] = df[col].apply(lambda x: clean_and_map(x, CATEGORY_MAPPING))
                def extract_cat_code(v):
                    try:
                        return int(float(str(v).split('-')[0].strip()))
                    except:
                        return 0
                df['Category_Code'] = df[col].apply(extract_cat_code)
                
        return df
    except Exception as e:
        st.error(f"Error reading UDISE file: {e}")
        return None

@st.cache_data(ttl=30)
def load_cadre_data(file_path):
    actual_path = None
    for f in os.listdir('.'):
        if f.lower() == file_path.lower():
            actual_path = f
            break
    if not actual_path:
        return None
    try:
        df_raw = pd.read_excel(actual_path, header=[0, 1])
        new_cols = []
        for c in df_raw.columns:
            l0 = str(c[0]).strip()
            l1 = str(c[1]).strip()
            if 'UNNAMED' in l0.upper() or l0 == '' or 'NAN' in l0.upper():
                new_cols.append(l1)
            elif 'UNNAMED' in l1.upper() or l1 == '' or 'NAN' in l1.upper():
                new_cols.append(l0)
            else:
                new_cols.append(f"{l0} - {l1}")
        df_raw.columns = new_cols
        df_raw = df_raw.loc[:, ~df_raw.columns.duplicated()]
        return df_raw
    except Exception as e:
        try:
            df_fallback = pd.read_excel(actual_path)
            df_fallback = df_fallback.loc[:, ~df_fallback.columns.duplicated()]
            return df_fallback
        except:
            return None

@st.cache_data(ttl=30)
def load_mbu_data(file_path):
    actual_path = None
    for f in os.listdir('.'):
        if f.lower() == file_path.lower():
            actual_path = f
            break
    if not actual_path:
        return None
    try:
        df_m = pd.read_excel(actual_path, header=None)
        header_row_idx = 0
        for i in range(min(5, len(df_m))):
            row_vals = [str(x).upper() for x in df_m.iloc[i].values]
            if any('UDISE' in x for x in row_vals) or any('BLOCK' in x for x in row_vals):
                header_row_idx = i
                break
                
        df_m = pd.read_excel(actual_path, skiprows=header_row_idx)
        df_m.columns = [" ".join(str(c).replace('\xa0', ' ').replace('\n', ' ').split()).strip() for c in df_m.columns]
        
        mbu_b_col = next((c for c in df_m.columns if 'BLOCK' in c.upper() or 'MANDAL' in c.upper()), None)
        if mbu_b_col:
            df_m[mbu_b_col] = df_m[mbu_b_col].astype(str).str.strip()
            df_m = df_m[~df_m[mbu_b_col].str.match(r'^\(?\d+\)?$|^nan$', case=False)]
            df_m = df_m[df_m[mbu_b_col].str.len() > 2]

        def clean_num(series):
            return pd.to_numeric(series.astype(str).str.replace(',', '').str.strip(), errors='coerce').fillna(0)

        for col in df_m.columns:
            if 'MANAGEMENT' in col.upper():
                df_m['Management_Display'] = df_m[col].apply(lambda x: clean_and_map(x, MANAGEMENT_MAPPING))
                df_m['Merged_Management'] = df_m[col].apply(get_merged_mgmt)
            if 'CATEGORY' in col.upper():
                df_m['Category_Display'] = df_m[col].apply(lambda x: clean_and_map(x, CATEGORY_MAPPING))

        p5_15_col = next((c for c in df_m.columns if '5-15' in c or ('PENDING' in c.upper() and '5' in c)), None)
        p15_p_col = next((c for c in df_m.columns if '15 AND ABOVE' in c.upper() or '15+' in c or ('PENDING' in c.upper() and 'ABOVE' in c.upper())), None)
        
        val_5_15 = clean_num(df_m[p5_15_col]) if p5_15_col else 0
        val_15_p = clean_num(df_m[p15_p_col]) if p15_p_col else 0
        
        df_m['MBU_P_5_15'] = val_5_15
        df_m['MBU_P_15_PLUS'] = val_15_p
        df_m['Total_MBU_Pending'] = val_5_15 + val_15_p
        
        if df_m['Total_MBU_Pending'].sum() == 0:
            general_pend = next((c for c in df_m.columns if 'PENDING' in c.upper() and 'NOT' not in c.upper()), None)
            if general_pend:
                df_m['Total_MBU_Pending'] = clean_num(df_m[general_pend])

        return df_m
    except Exception as e:
        st.error(f"Error reading MBU file: {e}")
        return None

df = load_udise_data(EXCEL_FILE_PATH)
df_cadre = load_cadre_data(TEACHERS_FILE_PATH)
df_mbu = load_mbu_data(MBU_FILE_PATH)

tab1, tab2, tab3, tab4 = st.tabs([
    "🏫 School 360° & UDISE Reports",
    "🧑‍🏫 Teachers Directory & Retirement",
    "📊 Cadre Strength & Vacancy",
    "📄 CSE MIS Reports"
])

# ==========================================
# ------------ TAB 1: UDISE ---------------
# ==========================================
with tab1:
    if df is None:
        st.error(f"⚠️ File dorakaledhu: {EXCEL_FILE_PATH}")
    else:
        udise_col = next((c for c in df.columns if c.upper() == 'UDISE CODE' or 'UDISE' in c.upper()), None)
        school_col = next((c for c in df.columns if c.upper() == 'SCHOOL NAME'), None)
        block_col = next((c for c in df.columns if 'BLOCK NAME' in c.upper() or 'MANDAL' in c.upper()), None)
        
        tot_col = next((c for c in df.columns if c.upper() in ['GRAND TOTAL', 'TOTAL ENROLMENT', 'TOTAL ROLL', 'TOTAL']), None)
        boys_col = next((c for c in df.columns if c.upper() in ['TOTAL BOYS', 'BOYS TOTAL', 'BOYS']), None)
        girls_col = next((c for c in df.columns if c.upper() in ['TOTAL GIRLS', 'GIRLS TOTAL', 'GIRLS']), None)
        
        if not tot_col:
            candidate_totals = [c for c in df.columns if 'TOTAL' in c.upper() and '(' not in c]
            tot_col = candidate_totals[-1] if candidate_totals else None

        if tot_col:
            df[tot_col] = pd.to_numeric(df[tot_col], errors='coerce').fillna(0)
        if boys_col:
            df[boys_col] = pd.to_numeric(df[boys_col], errors='coerce').fillna(0)
        if girls_col:
            df[girls_col] = pd.to_numeric(df[girls_col], errors='coerce').fillna(0)

        mgmt_col = 'Management_Display' if 'Management_Display' in df.columns else next((c for c in df.columns if 'MANAGE' in c.upper()), None)
        cat_col = 'Category_Display' if 'Category_Display' in df.columns else next((c for c in df.columns if 'CATEG' in c.upper()), None)

        def get_pp_cols(gender):
            res = []
            for c in df.columns:
                cu = c.upper()
                if any(p in cu for p in ['PP', 'PRE', 'LKG', 'UKG', 'NURSERY']) and gender.upper() in cu:
                    res.append(c)
            return list(set(res))

        def get_class_cols(classes, gender):
            res = []
            for c in df.columns:
                cu = c.upper()
                if any(p in cu for p in ['PP', 'PRE', 'LKG', 'UKG', 'NURSERY']):
                    continue
                for cl in classes:
                    patterns = [f"C{cl}(", f"CLASS {cl}(", f"CLASS{cl}(", f" {cl}("]
                    if any(p in cu for p in patterns) and gender.upper() in cu:
                        res.append(c)
            return list(set(res))

        def calc_sum(dframe, col_list):
            if not col_list:
                return pd.Series(0, index=dframe.index)
            return dframe[col_list].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1)

        pp_b = get_pp_cols('BOY')
        pp_g = get_pp_cols('GIRL')
        c1_5_b = get_class_cols([1, 2, 3, 4, 5], 'BOY')
        c1_5_g = get_class_cols([1, 2, 3, 4, 5], 'GIRL')
        c6_8_b = get_class_cols([6, 7, 8], 'BOY')
        c6_8_g = get_class_cols([6, 7, 8], 'GIRL')
        c6_10_b = get_class_cols([6, 7, 8, 9, 10], 'BOY')
        c6_10_g = get_class_cols([6, 7, 8, 9, 10], 'GIRL')
        c11_12_b = get_class_cols([11, 12], 'BOY')
        c11_12_g = get_class_cols([11, 12], 'GIRL')

        cat_series = df['Category_Code'] if 'Category_Code' in df.columns else pd.Series(0, index=df.index)
        
        # 1. Pre-Primary (PP 1-3)
        df['PP_B'] = calc_sum(df, pp_b)
        df['PP_G'] = calc_sum(df, pp_g)
        df['PP_T'] = df['PP_B'] + df['PP_G']

        # 2. Strict Class 1-5 (Excluding PP)
        mask_1_5 = cat_series.isin([1, 2, 3, 6])
        df['P_1_5_B'] = np.where(mask_1_5, calc_sum(df, c1_5_b), 0)
        df['P_1_5_G'] = np.where(mask_1_5, calc_sum(df, c1_5_g), 0)
        df['P_1_5_T'] = df['P_1_5_B'] + df['P_1_5_G']

        # 3. 6-8 UP
        mask_6_8_up = cat_series.isin([2])
        df['UP_6_8_B'] = np.where(mask_6_8_up, calc_sum(df, c6_8_b), 0)
        df['UP_6_8_G'] = np.where(mask_6_8_up, calc_sum(df, c6_8_g), 0)
        df['UP_6_8_T'] = df['UP_6_8_B'] + df['UP_6_8_G']

        # 4. 6-10 HS
        mask_6_10_hs = cat_series.isin([3, 5, 6, 7])
        df['HS_6_10_B'] = np.where(mask_6_10_hs, calc_sum(df, c6_10_b), 0)
        df['HS_6_10_G'] = np.where(mask_6_10_hs, calc_sum(df, c6_10_g), 0)
        df['HS_6_10_T'] = df['HS_6_10_B'] + df['HS_6_10_G']

        # 5. 11-12 Col
        mask_11_12 = cat_series.isin([11, 3, 5])
        df['COL_11_12_B'] = np.where(mask_11_12, calc_sum(df, c11_12_b), 0)
        df['COL_11_12_G'] = np.where(mask_11_12, calc_sum(df, c11_12_g), 0)
        df['COL_11_12_T'] = df['COL_11_12_B'] + df['COL_11_12_G']

        subtab1, subtab2, subtab3, subtab4 = st.tabs([
            "🔍 School 360° Search", 
            "📊 Mandal-wise Abstract", 
            "📑 Custom Reports & Excel Export",
            "⏳ Mandatory Biometric (MBU) Pending"
        ])

        with subtab1:
            st.subheader("🏫 Individual School 360° Profile")
            
            # --- DISTRICT SUMMARY CARDS ---
            d_s1, d_s2, d_s3, d_s4 = st.columns(4)
            d_s1.metric("District Total Schools 🏫", f"{len(df):,}")
            d_s2.metric("Total District Roll 👥", f"{int(df[tot_col].sum()):,}")
            d_s3.metric("District Boys 👦", f"{int(df[boys_col].sum()):,}")
            d_s4.metric("District Girls 👧", f"{int(df[girls_col].sum()):,}")
            st.markdown("---")

            col_search1, col_search2 = st.columns([3, 1])
            with col_search1:
                search_code = st.text_input("Enter 11 Digit UDISE Code:", value="28153500204")
            with col_search2:
                st.write("")
                st.write("")
                search_btn = st.button("Search Profile", use_container_width=True)

            if search_code and udise_col:
                matched = df[df[udise_col].astype(str).str.contains(str(search_code).strip(), na=False)]
                if not matched.empty:
                    row = matched.iloc[0]
                    school_name = row[school_col] if school_col else "School Name"
                    mandal_name = row[block_col] if block_col else "N/A"
                    mgmt_val = row[mgmt_col] if mgmt_col else "N/A"
                    cat_val = row[cat_col] if cat_col else "N/A"
                    
                    st.success(f"### 🏫 {school_name}")
                    st.info(f"**UDISE:** {row[udise_col]} | **Mandal (Block):** {mandal_name} | **Management:** {mgmt_val} | **Category:** {cat_val}")
                    
                    m1, m2, m3 = st.columns(3)
                    tot_val = int(row[tot_col]) if tot_col else 0
                    b_val = int(row[boys_col]) if boys_col else 0
                    g_val = int(row[girls_col]) if girls_col else 0
                    m1.metric("Grand Total Enrolment", f"{tot_val:,}")
                    m2.metric("Total Boys 👦", f"{b_val:,}")
                    m3.metric("Total Girls 👧", f"{g_val:,}")

                    st.markdown("---")
                    st.markdown("#### 📋 Class-wise Enrolment Breakdown")
                    
                    all_cols = list(df.columns)
                    class_keys = []
                    for c in all_cols:
                        if '(' in c and ')' in c:
                            key = c.split('(')[0].strip()
                            if key not in class_keys and any(tag in c.upper() for tag in ['BOY', 'GIRL', 'TOTAL', 'TRANS']):
                                class_keys.append(key)

                    class_data = []
                    for ck in class_keys:
                        b_col_k = next((c for c in all_cols if c.startswith(ck) and 'BOY' in c.upper()), None)
                        g_col_k = next((c for c in all_cols if c.startswith(ck) and 'GIRL' in c.upper()), None)
                        t_col_k = next((c for c in all_cols if c.startswith(ck) and 'TOTAL' in c.upper()), None)
                        
                        b_num = int(pd.to_numeric(row[b_col_k], errors='coerce')) if b_col_k else 0
                        g_num = int(pd.to_numeric(row[g_col_k], errors='coerce')) if g_col_k else 0
                        t_num = int(pd.to_numeric(row[t_col_k], errors='coerce')) if t_col_k else (b_num + g_num)
                        
                        if t_num > 0 or b_num > 0 or g_num > 0:
                            class_data.append({
                                "Class": ck,
                                "Boys 👦": b_num,
                                "Girls 👧": g_num,
                                "Total Enrolment": t_num
                            })

                    if class_data:
                        cdf = pd.DataFrame(class_data)
                        cdf_with_total = append_total_row(cdf, label_col="Class", total_label="TOTAL")
                        col_t1, col_t2 = st.columns([3, 2])
                        with col_t1:
                            st.dataframe(cdf_with_total, use_container_width=True, hide_index=True)
                            render_print_button(cdf_with_total, report_title=f"{school_name} - CLASS-WISE ENROLMENT", subtitle=f"UDISE: {row[udise_col]} | Mandal: {mandal_name}")
                        with col_t2:
                            st.bar_chart(cdf.set_index("Class")[["Boys 👦", "Girls 👧"]])
                    else:
                        st.info("Ee paatasaalaku sambandhinchina tharagathula vivaraalu labhinchaledhu.")
                else:
                    st.warning("Ee UDISE code tho record kanabada ledhu.")

        # --- SUBTAB 2: MANDAL-WISE ABSTRACT & DISTRICT MANAGEMENT SUMMARY ---
        with subtab2:
            st.subheader("📊 Mandal-wise Stage & Management Enrolment Abstract")

            df_clean = df[~df[block_col].astype(str).str.match(r'^\(?\d+\)?$|^nan$', case=False)].copy()
            df_clean = df_clean[df_clean[block_col].astype(str).str.len() > 2]

            # --- DISTRICT SUMMARY CARDS FOR STAGES ---
            ds_1, ds_2, ds_3, ds_4, ds_5 = st.columns(5)
            ds_1.metric("PP (1-3) Enrolment", f"{int(df_clean['PP_T'].sum()):,}")
            ds_2.metric("Class 1-5 Enrolment", f"{int(df_clean['P_1_5_T'].sum()):,}")
            ds_3.metric("Class 6-8 UP Enrolment", f"{int(df_clean['UP_6_8_T'].sum()):,}")
            ds_4.metric("Class 6-10 HS Enrolment", f"{int(df_clean['HS_6_10_T'].sum()):,}")
            ds_5.metric("Class 11-12 Col Enrolment", f"{int(df_clean['COL_11_12_T'].sum()):,}")
            st.markdown("---")

            mgmt_choice = st.selectbox(
                "Select Management to View Stage Breakdown (PP 1-3, 1-5, 6-8, 6-10, 11-12):",
                ["ALL MANAGEMENTS (Total District)", "STATE GOVT Only", "AIDED Only", "PRIVATE Only", "COMPARATIVE VIEW (Govt vs Aided vs Pvt Total Roll)"]
            )

            # TOP MANDAL-WISE TABLE (UNDISTURBED)
            if mgmt_choice == "COMPARATIVE VIEW (Govt vs Aided vs Pvt Total Roll)":
                piv_s = df_clean.pivot_table(index=block_col, columns='Merged_Management', values=udise_col, aggfunc='count', fill_value=0)
                piv_b = df_clean.pivot_table(index=block_col, columns='Merged_Management', values=boys_col, aggfunc='sum', fill_value=0)
                piv_g = df_clean.pivot_table(index=block_col, columns='Merged_Management', values=girls_col, aggfunc='sum', fill_value=0)
                piv_r = df_clean.pivot_table(index=block_col, columns='Merged_Management', values=tot_col, aggfunc='sum', fill_value=0)

                records = []
                for m_name in sorted(df_clean[block_col].unique()):
                    sg_s = int(piv_s.loc[m_name, 'STATE GOVT']) if 'STATE GOVT' in piv_s.columns and m_name in piv_s.index else 0
                    sg_b = int(piv_b.loc[m_name, 'STATE GOVT']) if 'STATE GOVT' in piv_b.columns and m_name in piv_b.index else 0
                    sg_g = int(piv_g.loc[m_name, 'STATE GOVT']) if 'STATE GOVT' in piv_g.columns and m_name in piv_g.index else 0
                    sg_r = int(piv_r.loc[m_name, 'STATE GOVT']) if 'STATE GOVT' in piv_r.columns and m_name in piv_r.index else 0

                    ai_s = int(piv_s.loc[m_name, 'AIDED']) if 'AIDED' in piv_s.columns and m_name in piv_s.index else 0
                    ai_b = int(piv_b.loc[m_name, 'AIDED']) if 'AIDED' in piv_b.columns and m_name in piv_b.index else 0
                    ai_g = int(piv_g.loc[m_name, 'AIDED']) if 'AIDED' in piv_g.columns and m_name in piv_g.index else 0
                    ai_r = int(piv_r.loc[m_name, 'AIDED']) if 'AIDED' in piv_r.columns and m_name in piv_r.index else 0

                    pr_s = int(piv_s.loc[m_name, 'PRIVATE']) if 'PRIVATE' in piv_s.columns and m_name in piv_s.index else 0
                    pr_b = int(piv_b.loc[m_name, 'PRIVATE']) if 'PRIVATE' in piv_b.columns and m_name in piv_b.index else 0
                    pr_g = int(piv_g.loc[m_name, 'PRIVATE']) if 'PRIVATE' in piv_g.columns and m_name in piv_g.index else 0
                    pr_r = int(piv_r.loc[m_name, 'PRIVATE']) if 'PRIVATE' in piv_r.columns and m_name in piv_r.index else 0

                    records.append({
                        "Mandal (Block)": m_name,
                        "Govt Sch": sg_s, "Govt Boys": sg_b, "Govt Girls": sg_g, "Govt Total": sg_r,
                        "Aided Sch": ai_s, "Aided Boys": ai_b, "Aided Girls": ai_g, "Aided Total": ai_r,
                        "Pvt Sch": pr_s, "Pvt Boys": pr_b, "Pvt Girls": pr_g, "Pvt Total": pr_r,
                        "Total Sch": sg_s + ai_s + pr_s, 
                        "Grand Total Boys": sg_b + ai_b + pr_b, 
                        "Grand Total Girls": sg_g + ai_g + pr_g, 
                        "Grand Total Roll": sg_r + ai_r + pr_r
                    })

                res_df = pd.DataFrame(records)
                res_df_total = append_total_row(res_df, label_col="Mandal (Block)", total_label="DISTRICT TOTAL")
                st.dataframe(res_df_total, use_container_width=True, hide_index=True)

                col_btn1, col_btn2 = st.columns([1, 1])
                with col_btn1:
                    m_buf = io.BytesIO()
                    with pd.ExcelWriter(m_buf, engine='openpyxl') as writer:
                        res_df_total.to_excel(writer, index=False, sheet_name='Comparative_Abstract')
                    st.download_button("📥 Download Comparative Abstract (Excel)", data=m_buf.getvalue(), file_name="Mandal_Comparative_Management_Abstract.xlsx", use_container_width=True)
                with col_btn2:
                    render_print_button(res_df_total, report_title="MANDAL-WISE MANAGEMENT COMPARATIVE ABSTRACT", subtitle="Govt vs Aided vs Private")

            else:
                if mgmt_choice == "STATE GOVT Only":
                    df_target = df_clean[df_clean['Merged_Management'] == 'STATE GOVT']
                    label_sub = "STATE GOVT"
                elif mgmt_choice == "AIDED Only":
                    df_target = df_clean[df_clean['Merged_Management'] == 'AIDED']
                    label_sub = "AIDED"
                elif mgmt_choice == "PRIVATE Only":
                    df_target = df_clean[df_clean['Merged_Management'] == 'PRIVATE']
                    label_sub = "PRIVATE"
                else:
                    df_target = df_clean
                    label_sub = "ALL MANAGEMENTS"

                agg_dict = {
                    udise_col: 'count', tot_col: 'sum',
                    'PP_B': 'sum', 'PP_G': 'sum', 'PP_T': 'sum',
                    'P_1_5_B': 'sum', 'P_1_5_G': 'sum', 'P_1_5_T': 'sum',
                    'UP_6_8_B': 'sum', 'UP_6_8_G': 'sum', 'UP_6_8_T': 'sum',
                    'HS_6_10_B': 'sum', 'HS_6_10_G': 'sum', 'HS_6_10_T': 'sum',
                    'COL_11_12_B': 'sum', 'COL_11_12_G': 'sum', 'COL_11_12_T': 'sum'
                }

                mandal_res = df_target.groupby(block_col).agg(agg_dict).reset_index()
                mandal_res = mandal_res.rename(columns={
                    block_col: 'Mandal (Block)',
                    udise_col: 'Schools',
                    tot_col: 'Grand Total Roll',
                    'PP_B': 'PP (1-3) Boys', 'PP_G': 'PP (1-3) Girls', 'PP_T': 'PP (1-3) Total',
                    'P_1_5_B': '1-5 Boys', 'P_1_5_G': '1-5 Girls', 'P_1_5_T': '1-5 Total',
                    'UP_6_8_B': '6-8 UP Boys', 'UP_6_8_G': '6-8 UP Girls', 'UP_6_8_T': '6-8 UP Total',
                    'HS_6_10_B': '6-10 HS Boys', 'HS_6_10_G': '6-10 HS Girls', 'HS_6_10_T': '6-10 HS Total',
                    'COL_11_12_B': '11-12 Col Boys', 'COL_11_12_G': '11-12 Col Girls', 'COL_11_12_T': '11-12 Col Total'
                })

                stage_cols = [
                    'Mandal (Block)', 'Schools', 'Grand Total Roll',
                    'PP (1-3) Boys', 'PP (1-3) Girls', 'PP (1-3) Total',
                    '1-5 Boys', '1-5 Girls', '1-5 Total',
                    '6-8 UP Boys', '6-8 UP Girls', '6-8 UP Total',
                    '6-10 HS Boys', '6-10 HS Girls', '6-10 HS Total',
                    '11-12 Col Boys', '11-12 Col Girls', '11-12 Col Total'
                ]
                display_df = mandal_res[stage_cols]
                display_df_total = append_total_row(display_df, label_col='Mandal (Block)', total_label='DISTRICT TOTAL')
                st.dataframe(display_df_total, use_container_width=True, hide_index=True)

                col_btn1, col_btn2 = st.columns([1, 1])
                with col_btn1:
                    m_buf = io.BytesIO()
                    with pd.ExcelWriter(m_buf, engine='openpyxl') as writer:
                        display_df_total.to_excel(writer, index=False, sheet_name=label_sub[:31])
                    st.download_button(
                        f"📥 Download {label_sub} Stage Abstract Excel", 
                        data=m_buf.getvalue(), 
                        file_name=f"Mandal_{label_sub}_Stage_Wise_Enrolment.xlsx", 
                        use_container_width=True
                    )
                with col_btn2:
                    render_print_button(
                        display_df_total, 
                        report_title=f"MANDAL-WISE STAGE-WISE ABSTRACT - {label_sub}", 
                        subtitle="PP (1-3), 1-5 (Pr/UP/HS), 6-8 (UP), 6-10 (HS), 11-12 (Colleges)"
                    )

            # --- KOTHTHA REPORT: DISTRICT MANAGEMENT-WISE SUMMARY (GOVT / AIDED / PVT) ---
            st.markdown("---")
            st.markdown("#### 🏛️ District Management-wise Stage Abstract (Govt vs Aided vs Pvt)")
            st.caption("District-level consolidation across Pre-Primary, Primary, Upper Primary, High School & College Stages")

            mgmt_agg_dict = {
                udise_col: 'count', tot_col: 'sum',
                'PP_B': 'sum', 'PP_G': 'sum', 'PP_T': 'sum',
                'P_1_5_B': 'sum', 'P_1_5_G': 'sum', 'P_1_5_T': 'sum',
                'UP_6_8_B': 'sum', 'UP_6_8_G': 'sum', 'UP_6_8_T': 'sum',
                'HS_6_10_B': 'sum', 'HS_6_10_G': 'sum', 'HS_6_10_T': 'sum',
                'COL_11_12_B': 'sum', 'COL_11_12_G': 'sum', 'COL_11_12_T': 'sum'
            }

            dist_mgmt_summary = df_clean.groupby('Merged_Management').agg(mgmt_agg_dict).reset_index()
            dist_mgmt_summary = dist_mgmt_summary.rename(columns={
                'Merged_Management': 'Management',
                udise_col: 'Schools',
                tot_col: 'Grand Total Roll',
                'PP_B': 'PP (1-3) Boys', 'PP_G': 'PP (1-3) Girls', 'PP_T': 'PP (1-3) Total',
                'P_1_5_B': '1-5 Boys', 'P_1_5_G': '1-5 Girls', 'P_1_5_T': '1-5 Total',
                'UP_6_8_B': '6-8 UP Boys', 'UP_6_8_G': '6-8 UP Girls', 'UP_6_8_T': '6-8 UP Total',
                'HS_6_10_B': '6-10 HS Boys', 'HS_6_10_G': '6-10 HS Girls', 'HS_6_10_T': '6-10 HS Total',
                'COL_11_12_B': '11-12 Col Boys', 'COL_11_12_G': '11-12 Col Girls', 'COL_11_12_T': '11-12 Col Total'
            })

            order_map = {'STATE GOVT': 1, 'AIDED': 2, 'PRIVATE': 3}
            dist_mgmt_summary['order'] = dist_mgmt_summary['Management'].map(lambda x: order_map.get(x, 4))
            dist_mgmt_summary = dist_mgmt_summary.sort_values(by='order').drop(columns=['order'])

            dist_mgmt_cols = [
                'Management', 'Schools', 'Grand Total Roll',
                'PP (1-3) Boys', 'PP (1-3) Girls', 'PP (1-3) Total',
                '1-5 Boys', '1-5 Girls', '1-5 Total',
                '6-8 UP Boys', '6-8 UP Girls', '6-8 UP Total',
                '6-10 HS Boys', '6-10 HS Girls', '6-10 HS Total',
                '11-12 Col Boys', '11-12 Col Girls', '11-12 Col Total'
            ]
            dist_mgmt_display = dist_mgmt_summary[dist_mgmt_cols]
            dist_mgmt_total = append_total_row(dist_mgmt_display, label_col='Management', total_label='DISTRICT TOTAL')
            
            st.dataframe(dist_mgmt_total, use_container_width=True, hide_index=True)

            col_dm1, col_dm2 = st.columns([1, 1])
            with col_dm1:
                dm_buf = io.BytesIO()
                with pd.ExcelWriter(dm_buf, engine='openpyxl') as writer:
                    dist_mgmt_total.to_excel(writer, index=False, sheet_name='District_Management_Stage')
                st.download_button(
                    "📥 Download District Management Abstract Excel", 
                    data=dm_buf.getvalue(), 
                    file_name="District_Management_Wise_Stage_Abstract.xlsx", 
                    use_container_width=True
                )
            with col_dm2:
                render_print_button(
                    dist_mgmt_total, 
                    report_title="DISTRICT MANAGEMENT-WISE STAGE ENROLMENT ABSTRACT", 
                    subtitle="State Govt vs Aided vs Private (PP 1-3, 1-5, 6-8 UP, 6-10 HS, 11-12 Col)"
                )

        # --- SUBTAB 3: CUSTOM REPORTS (CLEANED UP TO HIDE INTERNAL CALC COLUMNS) ---
        with subtab3:
            st.subheader("📑 Custom Reports & Excel Export")
            
            f1, f2 = st.columns(2)
            with f1:
                mandal_list = sorted([str(m) for m in df[block_col].dropna().unique() if len(str(m).strip()) > 2 and not str(m).strip().startswith('(')]) if block_col else []
                sel_mandals = st.multiselect("Select Mandal(s):", mandal_list, default=mandal_list)
            with f2:
                mgmt_list = sorted(list(df[mgmt_col].dropna().unique())) if mgmt_col else []
                sel_mgmt = st.multiselect("Select Management:", mgmt_list, default=mgmt_list)

            cols_to_hide = ['PP_B', 'PP_G', 'PP_T', 'P_1_5_B', 'P_1_5_G', 'P_1_5_T', 'UP_6_8_B', 'UP_6_8_G', 'UP_6_8_T', 
                            'HS_6_10_B', 'HS_6_10_G', 'HS_6_10_T', 'COL_11_12_B', 'COL_11_12_G', 'COL_11_12_T', 'Category_Code']
            clean_display_df = df.drop(columns=[c for c in cols_to_hide if c in df.columns])

            filtered_df = clean_display_df.copy()
            if block_col and sel_mandals:
                filtered_df = filtered_df[filtered_df[block_col].isin(sel_mandals)]
            if mgmt_col and sel_mgmt:
                filtered_df = filtered_df[filtered_df[mgmt_col].isin(sel_mgmt)]

            # --- SUMMARY METRICS FOR FILTERED DATA ---
            c_m1, c_m2, c_m3, c_m4 = st.columns(4)
            c_m1.metric("Selected Schools 🏫", f"{len(filtered_df):,}")
            c_m2.metric("Total Selected Roll 👥", f"{int(filtered_df[tot_col].sum()):,}")
            c_m3.metric("Selected Boys 👦", f"{int(filtered_df[boys_col].sum()):,}")
            c_m4.metric("Selected Girls 👧", f"{int(filtered_df[girls_col].sum()):,}")
            st.markdown("---")

            st.dataframe(filtered_df, use_container_width=True)

            col_cf1, col_cf2 = st.columns([1, 1])
            with col_cf1:
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    filtered_df.to_excel(writer, index=False, sheet_name='Filtered_Report')
                st.download_button("📥 Download Filtered Report as Excel", data=buffer.getvalue(), file_name="Filtered_UDISE_Report.xlsx", use_container_width=True)
            with col_cf2:
                render_print_button(filtered_df.head(100), report_title="CUSTOM UDISE REPORT", subtitle=f"Total Schools: {len(filtered_df)}")

        # --- SUBTAB 4: MANDATORY BIOMETRIC PENDING (MERGED MANAGEMENT INTEGRATED) ---
        with subtab4:
            st.subheader("⏳ Mandatory Biometric Update (MBU) Pending Analysis")
            if df_mbu is None:
                st.warning(f"⚠️ '{MBU_FILE_PATH}' file GitHub lo load kaaledhu. File upload aindo ledho chudandi.")
            else:
                mbu_block_col = next((c for c in df_mbu.columns if 'BLOCK' in c.upper() or 'MANDAL' in c.upper()), None)
                mbu_mgmt_col = 'Management_Display' if 'Management_Display' in df_mbu.columns else next((c for c in df_mbu.columns if 'MANAGE' in c.upper()), None)
                mbu_udise_col = next((c for c in df_mbu.columns if 'UDISE' in c.upper()), None)
                mbu_school_col = next((c for c in df_mbu.columns if 'SCHOOL' in c.upper() and 'CATEGORY' not in c.upper() and 'MANAGEMENT' not in c.upper()), None)
                
                tot_stu_col = next((c for c in df_mbu.columns if 'TOTAL STUDENT' in c.upper() or 'TOTAL' in c.upper()), None)
                passed_col = next((c for c in df_mbu.columns if 'PASSED' in c.upper()), None)
                failed_col = next((c for c in df_mbu.columns if 'FAILED' in c.upper()), None)
                
                tot_mbu_pend = int(df_mbu['Total_MBU_Pending'].sum())
                schools_with_pend = int((df_mbu['Total_MBU_Pending'] > 0).sum())
                
                # --- DISTRICT SUMMARY CARDS FOR MBU ---
                mb_m1, mb_m2, mb_m3, mb_m4 = st.columns(4)
                mb_m1.metric("Total District MBU Pending ⚠️", f"{tot_mbu_pend:,}")
                mb_m2.metric("Schools with MBU Pending 🏫", f"{schools_with_pend:,}")
                mb_m3.metric("District Pending (5-15) 👦", f"{int(df_mbu['MBU_P_5_15'].sum()):,}")
                mb_m4.metric("District Pending (15+) 🧑", f"{int(df_mbu['MBU_P_15_PLUS'].sum()):,}")
                st.markdown("---")

                mbu_view1, mbu_view2 = st.tabs([
                    "📊 Mandal-wise Merged Management MBU Abstract", 
                    "🏫 Mandal-wise School Detailed List"
                ])

                with mbu_view1:
                    st.markdown("##### 📌 Mandal-wise: Govt, Aided & Private MBU Pending (5-15 & 15+)")
                    if mbu_block_col and 'Merged_Management' in df_mbu.columns:
                        mbu_clean = df_mbu.copy()
                        
                        piv_5_15 = mbu_clean.pivot_table(index=mbu_block_col, columns='Merged_Management', values='MBU_P_5_15', aggfunc='sum', fill_value=0)
                        piv_15_p = mbu_clean.pivot_table(index=mbu_block_col, columns='Merged_Management', values='MBU_P_15_PLUS', aggfunc='sum', fill_value=0)

                        mbu_records = []
                        for m_name in sorted(mbu_clean[mbu_block_col].unique()):
                            g_5 = int(piv_5_15.loc[m_name, 'STATE GOVT']) if 'STATE GOVT' in piv_5_15.columns and m_name in piv_5_15.index else 0
                            g_15 = int(piv_15_p.loc[m_name, 'STATE GOVT']) if 'STATE GOVT' in piv_15_p.columns and m_name in piv_15_p.index else 0
                            g_tot = g_5 + g_15

                            a_5 = int(piv_5_15.loc[m_name, 'AIDED']) if 'AIDED' in piv_5_15.columns and m_name in piv_5_15.index else 0
                            a_15 = int(piv_15_p.loc[m_name, 'AIDED']) if 'AIDED' in piv_15_p.columns and m_name in piv_15_p.index else 0
                            a_tot = a_5 + a_15

                            p_5 = int(piv_5_15.loc[m_name, 'PRIVATE']) if 'PRIVATE' in piv_5_15.columns and m_name in piv_5_15.index else 0
                            p_15 = int(piv_15_p.loc[m_name, 'PRIVATE']) if 'PRIVATE' in piv_15_p.columns and m_name in piv_15_p.index else 0
                            p_tot = p_5 + p_15

                            tot_5 = g_5 + a_5 + p_5
                            tot_15 = g_15 + a_15 + p_15
                            grand_pend = tot_5 + tot_15

                            mbu_records.append({
                                "Mandal (Block)": m_name,
                                "Govt (5-15)": g_5, "Govt (15+)": g_15, "Govt Total": g_tot,
                                "Aided (5-15)": a_5, "Aided (15+)": a_15, "Aided Total": a_tot,
                                "Pvt (5-15)": p_5, "Pvt (15+)": p_15, "Pvt Total": p_tot,
                                "Total (5-15)": tot_5, "Total (15+)": tot_15, "Grand Total Pending": grand_pend
                            })

                        mbu_mgmt_df = pd.DataFrame(mbu_records)
                        mbu_mgmt_df_total = append_total_row(mbu_mgmt_df, label_col="Mandal (Block)", total_label="DISTRICT TOTAL")
                        st.dataframe(mbu_mgmt_df_total, use_container_width=True, hide_index=True)

                        col_pb1, col_pb2 = st.columns([1, 1])
                        with col_pb1:
                            p_buf = io.BytesIO()
                            with pd.ExcelWriter(p_buf, engine='openpyxl') as writer:
                                mbu_mgmt_df_total.to_excel(writer, index=False, sheet_name='MBU_Management_Matrix')
                            st.download_button(
                                "📥 Download MBU Merged Management Matrix as Excel", 
                                data=p_buf.getvalue(), 
                                file_name="Mandal_Management_MBU_Pending_5_15.xlsx",
                                use_container_width=True
                            )
                        with col_pb2:
                            render_print_button(
                                mbu_mgmt_df_total, 
                                report_title="MANDAL-WISE MERGED MANAGEMENT MBU PENDING ABSTRACT", 
                                subtitle="Govt vs Aided vs Private (5-15 & 15+ Pending)"
                            )

                with mbu_view2:
                    st.markdown("##### 🔍 School-wise Pending Details by Mandal")
                    if mbu_block_col:
                        m_list_mbu = sorted([
                            str(m) for m in df_mbu[mbu_block_col].dropna().unique() 
                            if len(str(m).strip()) >= 3 and not str(m).strip().startswith('(')
                        ])
                        sel_mbu_mandal = st.selectbox("Select Mandal:", m_list_mbu, key="mbu_mandal_select")
                        
                        m_filter_df = df_mbu[df_mbu[mbu_block_col] == sel_mbu_mandal].copy()
                        show_only_pending = st.checkbox("Show only Schools with Pending > 0", value=True)
                        if show_only_pending:
                            m_filter_df = m_filter_df[m_filter_df['Total_MBU_Pending'] > 0]
                        
                        display_cols = []
                        rename_disp = {}
                        if mbu_udise_col:
                            display_cols.append(mbu_udise_col)
                            rename_disp[mbu_udise_col] = 'UDISE Code'
                        if mbu_school_col:
                            display_cols.append(mbu_school_col)
                            rename_disp[mbu_school_col] = 'School Name'
                        if mbu_mgmt_col:
                            display_cols.append(mbu_mgmt_col)
                            rename_disp[mbu_mgmt_col] = 'Management'
                        if tot_stu_col:
                            display_cols.append(tot_stu_col)
                            rename_disp[tot_stu_col] = 'Total Students'
                        if passed_col:
                            display_cols.append(passed_col)
                            rename_disp[passed_col] = 'Aadhaar Verified (Pass)'
                        if failed_col:
                            display_cols.append(failed_col)
                            rename_disp[failed_col] = 'Aadhaar Failed'
                        if 'MBU_P_5_15' in m_filter_df.columns:
                            display_cols.append('MBU_P_5_15')
                            rename_disp['MBU_P_5_15'] = 'MBU Pending (5-15)'
                        if 'MBU_P_15_PLUS' in m_filter_df.columns:
                            display_cols.append('MBU_P_15_PLUS')
                            rename_disp['MBU_P_15_PLUS'] = 'MBU Pending (15+)'
                        
                        display_cols.append('Total_MBU_Pending')
                        rename_disp['Total_MBU_Pending'] = 'Grand Total Pending'

                        final_disp_cols = [c for c in display_cols if c in m_filter_df.columns]
                        sch_disp_df = m_filter_df[final_disp_cols].rename(columns=rename_disp)
                        
                        label_c = 'School Name' if 'School Name' in sch_disp_df.columns else sch_disp_df.columns[0]
                        sch_disp_with_total = append_total_row(sch_disp_df, label_col=label_c, total_label='MANDAL TOTAL')
                        
                        st.dataframe(sch_disp_with_total, use_container_width=True, hide_index=True)
                        
                        col_sb1, col_sb2 = st.columns([1, 1])
                        with col_sb1:
                            s_buf = io.BytesIO()
                            with pd.ExcelWriter(s_buf, engine='openpyxl') as writer:
                                sch_disp_with_total.to_excel(writer, index=False, sheet_name=str(sel_mbu_mandal)[:31])
                            st.download_button(
                                f"📥 Download {sel_mbu_mandal} MBU School Report (Excel)", 
                                data=s_buf.getvalue(), 
                                file_name=f"{sel_mbu_mandal}_MBU_Pending_Report.xlsx",
                                use_container_width=True
                            )
                        with col_sb2:
                            render_print_button(sch_disp_with_total, report_title=f"{sel_mbu_mandal} MANDAL - SCHOOL-WISE MBU PENDING REPORT", subtitle="School-wise Biometric Pending Details")

# ==========================================
# ------------ TAB 2: TEACHERS ------------
# ==========================================
with tab2:
    st.info("🧑‍🏫 Teachers Directory & Retirement Tracker - Module Coming Soon")

# ==========================================
# ------------ TAB 3: CADRE & VACANCY -----
# ==========================================
with tab3:
    st.subheader("📊 Cadre Strength, Working & Vacancy Analysis")
    if df_cadre is None:
        st.warning(f"⚠️ '{TEACHERS_FILE_PATH}' file load kaaledhu. File upload aindo ledho chudandi.")
    else:
        c_udise = next((c for c in df_cadre.columns if 'UDISE' in str(c).upper()), None)
        c_school = next((c for c in df_cadre.columns if any(k in str(c).upper() for k in ['HS/UPS_NAME', 'SCHOOL', 'NAME'])), None)
        c_mandal = next((c for c in df_cadre.columns if 'MANDAL' in str(c).upper()), None)
        
        sanc_cols = [c for c in df_cadre.columns if 'SANCTIONED' in str(c).upper() and 'TOTAL' not in str(c).upper()]
        work_cols = [c for c in df_cadre.columns if 'WORKING' in str(c).upper() and 'TOTAL' not in str(c).upper() and 'MTS' not in str(c).upper()]
        vac_cols  = [c for c in df_cadre.columns if 'VACANT' in str(c).upper() and 'TOTAL' not in str(c).upper() and 'MTS' not in str(c).upper()]
        
        tot_sanc_col = next((c for c in df_cadre.columns if 'SANCTIONED' in str(c).upper() and 'TOTAL' in str(c).upper()), None)
        tot_work_col = next((c for c in df_cadre.columns if 'WORKING' in str(c).upper() and 'TOTAL' in str(c).upper()), None)
        tot_vac_col  = next((c for c in df_cadre.columns if 'VACANT' in str(c).upper() and 'TOTAL' in str(c).upper()), None)

        def normalize_cadre_name(name):
            clean = str(name).strip()
            for kw in ['No of Posts Sanctioned', 'No of Posts Working', 'No of Posts Vacant', 'Sanctioned', 'Working', 'Vacant']:
                if kw.lower() in clean.lower():
                    parts = clean.split(' - ')
                    if len(parts) > 1:
                        clean = parts[-1].strip()
                    else:
                        clean = clean.replace(kw, '').strip(' -:')
            clean = clean.replace('01.09.2026', '').strip(' -:')
            if clean.lower().startswith('gr ii hm') or clean.lower().startswith('gr-ii hm'):
                return "Gr II HM"
            if 'SGT' in str(name).upper() and 'TELUGU' in str(name).upper():
                return "SGT - TELUGU"
            if 'SGT' in str(name).upper() and 'URDU' in str(name).upper():
                return "SGT URDU"
            return clean

        c_tab1, c_tab2, c_tab3 = st.tabs([
            "🔍 School Cadre Profile", 
            "📌 Mandal & Cadre Vacancies", 
            "📑 Mandal Cadre Summary (S/W/V)"
        ])

        with c_tab1:
            st.markdown("#### 🏫 Individual School Cadre Strength")
            cadre_search = st.text_input("Enter UDISE Code or School Name:", value="28153500204", key="cadre_srch_input")
            
            matched_cadre = pd.DataFrame()
            if cadre_search and c_udise:
                udise_series = df_cadre[c_udise]
                if isinstance(udise_series, pd.DataFrame):
                    udise_series = udise_series.iloc[:, 0]
                matched_cadre = df_cadre[udise_series.astype(str).str.contains(str(cadre_search).strip(), na=False)]
            
            if not matched_cadre.empty:
                c_row = matched_cadre.iloc[0]
                s_name = c_row[c_school] if c_school else "School"
                m_name = c_row[c_mandal] if c_mandal else "N/A"
                
                st.success(f"### 🏫 {s_name} ({m_name} Mandal)")
                
                col_m1, col_m2, col_m3 = st.columns(3)
                v_sanc = int(pd.to_numeric(c_row[tot_sanc_col], errors='coerce')) if tot_sanc_col and pd.notnull(c_row[tot_sanc_col]) else 0
                v_work = int(pd.to_numeric(c_row[tot_work_col], errors='coerce')) if tot_work_col and pd.notnull(c_row[tot_work_col]) else 0
                v_vac  = int(pd.to_numeric(c_row[tot_vac_col], errors='coerce')) if tot_vac_col and pd.notnull(c_row[tot_vac_col]) else 0
                
                col_m1.metric("Sanctioned Posts", v_sanc)
                col_m2.metric("Working Staff 👥", v_work)
                col_m3.metric("Vacant Posts ⚠️", v_vac)
                
                st.markdown("---")
                st.markdown("##### 📋 Post-wise Breakup (Sanctioned vs Working vs Vacant)")
                
                post_list = []
                for idx, sc in enumerate(sanc_cols):
                    clean_post = normalize_cadre_name(sc)
                    wc = work_cols[idx] if idx < len(work_cols) else None
                    vc = vac_cols[idx] if idx < len(vac_cols) else None
                    
                    if not wc or clean_post.lower() not in normalize_cadre_name(wc).lower():
                        wc = next((c for c in work_cols if clean_post.lower() == normalize_cadre_name(c).lower()), wc)
                    if not vc or clean_post.lower() not in normalize_cadre_name(vc).lower():
                        vc = next((c for c in vac_cols if clean_post.lower() == normalize_cadre_name(vc).lower()), vc)

                    s_val = int(pd.to_numeric(c_row[sc], errors='coerce')) if pd.notnull(c_row[sc]) else 0
                    w_val = int(pd.to_numeric(c_row[wc], errors='coerce')) if wc and pd.notnull(c_row[wc]) else 0
                    vac_val = int(pd.to_numeric(c_row[vc], errors='coerce')) if vc and pd.notnull(c_row[vc]) else 0
                    
                    if s_val > 0 or w_val > 0 or vac_val > 0:
                        post_list.append({
                            "Designation / Cadre": clean_post,
                            "Sanctioned": s_val,
                            "Working": w_val,
                            "Vacant": vac_val
                        })
                
                if post_list:
                    p_df = pd.DataFrame(post_list)
                    p_df_with_total = append_total_row(p_df, label_col="Designation / Cadre", total_label="TOTAL")
                    st.dataframe(p_df_with_total, use_container_width=True, hide_index=True)
                    render_print_button(p_df_with_total, report_title=f"{s_name} - CADRE BREAKUP", subtitle=f"UDISE: {c_row[c_udise]} | Mandal: {m_name}")
                else:
                    st.info("Ee paatasaalaku sambandhinchina post-wise vivaraalu levu.")
            else:
                st.info("Paatasaala vivaraalu chudadaniki UDISE code enter cheyandi.")

        with c_tab2:
            st.markdown("#### 📌 Mandal-wise & Cadre-wise Vacancy Matrix")
            if c_mandal and vac_cols:
                df_vac = df_cadre.copy()
                col_rename_map = {}
                for vc in vac_cols:
                    df_vac[vc] = pd.to_numeric(df_vac[vc], errors='coerce').fillna(0)
                    col_rename_map[vc] = normalize_cadre_name(vc)
                
                mandal_cadre_vac = df_vac.groupby(c_mandal)[vac_cols].sum().reset_index()
                mandal_cadre_vac = mandal_cadre_vac.rename(columns=col_rename_map)
                mandal_cadre_vac = mandal_cadre_vac.rename(columns={c_mandal: 'Mandal'})
                
                cadre_cols_cleaned = [col_rename_map[vc] for vc in vac_cols]
                mandal_cadre_vac['Total Vacancies'] = mandal_cadre_vac[cadre_cols_cleaned].sum(axis=1)
                
                all_mandals = sorted(mandal_cadre_vac['Mandal'].dropna().unique())
                sel_m = st.multiselect("Filter by Mandal(s):", all_mandals, default=all_mandals, key="vac_mandal_filter")
                
                display_vac_df = mandal_cadre_vac[mandal_cadre_vac['Mandal'].isin(sel_m)] if sel_m else mandal_cadre_vac
                cols_order = ['Mandal', 'Total Vacancies'] + [c for c in cadre_cols_cleaned if c != 'Total Vacancies']
                display_vac_df = display_vac_df[cols_order]
                
                display_vac_with_total = append_total_row(display_vac_df, label_col='Mandal', total_label='DISTRICT TOTAL')
                st.dataframe(display_vac_with_total, use_container_width=True, hide_index=True)
                
                col_vm1, col_vm2 = st.columns([1, 1])
                with col_vm1:
                    vac_matrix_buf = io.BytesIO()
                    with pd.ExcelWriter(vac_matrix_buf, engine='openpyxl') as writer:
                        display_vac_with_total.to_excel(writer, index=False, sheet_name='Mandal_Cadre_Vacancies')
                    st.download_button(
                        "📥 Download Vacancy Matrix as Excel", 
                        data=vac_matrix_buf.getvalue(), 
                        file_name="Mandal_Cadre_Wise_Vacancies.xlsx",
                        use_container_width=True
                    )
                with col_vm2:
                    render_print_button(display_vac_with_total, report_title="MANDAL-WISE & CADRE-WISE VACANCY MATRIX", subtitle="West Godavari District")
            else:
                st.info("Khaaleelu emee record kaledhu.")

        with c_tab3:
            st.markdown("#### 📑 Mandal-wise Cadre-wise Status (Sanctioned, Working, Vacant)")
            if c_mandal and sanc_cols:
                df_cadre_work = df_cadre.copy()
                for c in sanc_cols + work_cols + vac_cols:
                    if c in df_cadre_work.columns:
                        df_cadre_work[c] = pd.to_numeric(df_cadre_work[c], errors='coerce').fillna(0)

                # --- DISTRICT SUMMARY FOR CADRE ---
                dc_1, dc_2, dc_3 = st.columns(3)
                dc_1.metric("District Total Sanctioned 🏛️", f"{int(df_cadre_work[tot_sanc_col].sum()):,}" if tot_sanc_col else "0")
                dc_2.metric("District Working Staff 👥", f"{int(df_cadre_work[tot_work_col].sum()):,}" if tot_work_col else "0")
                dc_3.metric("District Total Vacancies ⚠️", f"{int(df_cadre_work[tot_vac_col].sum()):,}" if tot_vac_col else "0")
                st.markdown("---")

                mandal_list_all = sorted(list(df_cadre_work[c_mandal].dropna().unique()))
                selected_mandal = st.selectbox("Select Mandal to view detailed Cadre breakdown:", mandal_list_all, key="mandal_cadre_detailed_select")

                m_df = df_cadre_work[df_cadre_work[c_mandal] == selected_mandal]
                
                cadre_breakdown = []
                for idx, sc in enumerate(sanc_cols):
                    p_name = normalize_cadre_name(sc)
                    wc = work_cols[idx] if idx < len(work_cols) else None
                    vc = vac_cols[idx] if idx < len(vac_cols) else None

                    if not wc or p_name.lower() not in normalize_cadre_name(wc).lower():
                        wc = next((c for c in work_cols if p_name.lower() == normalize_cadre_name(c).lower()), wc)
                    if not vc or p_name.lower() not in normalize_cadre_name(vc).lower():
                        vc = next((c for c in vac_cols if p_name.lower() == normalize_cadre_name(vc).lower()), vc)

                    s_sum = int(m_df[sc].sum()) if sc in m_df.columns else 0
                    w_sum = int(m_df[wc].sum()) if wc and wc in m_df.columns else 0
                    v_sum = int(m_df[vc].sum()) if vc and vc in m_df.columns else 0

                    if s_sum > 0 or w_sum > 0 or v_sum > 0:
                        cadre_breakdown.append({
                            "Cadre / Designation": p_name,
                            "Sanctioned Posts": s_sum,
                            "Working Staff": w_sum,
                            "Vacant Posts": v_sum
                        })

                if cadre_breakdown:
                    cb_df = pd.DataFrame(cadre_breakdown)
                    cm1, cm2, cm3 = st.columns(3)
                    cm1.metric(f"Total Sanctioned ({selected_mandal})", cb_df["Sanctioned Posts"].sum())
                    cm2.metric(f"Total Working ({selected_mandal})", cb_df["Working Staff"].sum())
                    cm3.metric(f"Total Vacant ({selected_mandal})", cb_df["Vacant Posts"].sum())

                    cb_df_with_total = append_total_row(cb_df, label_col="Cadre / Designation", total_label="TOTAL")
                    st.dataframe(cb_df_with_total, use_container_width=True, hide_index=True)

                    col_mb1, col_mb2 = st.columns([1, 1])
                    with col_mb1:
                        s_buf = io.BytesIO()
                        with pd.ExcelWriter(s_buf, engine='openpyxl') as writer:
                            cb_df_with_total.to_excel(writer, index=False, sheet_name=str(selected_mandal)[:31])
                        st.download_button(
                            f"📥 Download {selected_mandal} Cadre Report (Excel)", 
                            data=s_buf.getvalue(), 
                            file_name=f"{selected_mandal}_Cadre_Breakdown.xlsx",
                            use_container_width=True
                        )
                    with col_mb2:
                        render_print_button(cb_df_with_total, report_title=f"{selected_mandal} MANDAL - CADRE STRENGTH & VACANCY", subtitle="Sanctioned vs Working vs Vacant")

                st.markdown("---")
                st.markdown("##### 🌐 Full District: Mandal-wise & Cadre-wise Master Table")
                
                master_records = []
                for m_name_iter in mandal_list_all:
                    sub_m = df_cadre_work[df_cadre_work[c_mandal] == m_name_iter]
                    for idx, sc in enumerate(sanc_cols):
                        p_name = normalize_cadre_name(sc)
                        wc = work_cols[idx] if idx < len(work_cols) else None
                        vc = vac_cols[idx] if idx < len(vac_cols) else None
                        if not wc or p_name.lower() not in normalize_cadre_name(wc).lower():
                            wc = next((c for c in work_cols if p_name.lower() == normalize_cadre_name(c).lower()), wc)
                        if not vc or p_name.lower() not in normalize_cadre_name(vc).lower():
                            vc = next((c for c in vac_cols if p_name.lower() == normalize_cadre_name(vc).lower()), vc)

                        s_val = int(sub_m[sc].sum()) if sc in sub_m.columns else 0
                        w_val = int(sub_m[wc].sum()) if wc and wc in sub_m.columns else 0
                        v_val = int(sub_m[vc].sum()) if vc and vc in sub_m.columns else 0

                        if s_val > 0 or w_val > 0 or v_val > 0:
                            master_records.append({
                                "Mandal": m_name_iter,
                                "Cadre / Designation": p_name,
                                "Sanctioned": s_val,
                                "Working": w_val,
                                "Vacant": v_val
                            })

                if master_records:
                    master_df = pd.DataFrame(master_records)
                    master_df_with_total = append_total_row(master_df, label_col="Mandal", total_label="DISTRICT TOTAL")
                    st.dataframe(master_df_with_total, use_container_width=True, hide_index=True)

                    col_all1, col_all2 = st.columns([1, 1])
                    with col_all1:
                        master_buf = io.BytesIO()
                        with pd.ExcelWriter(master_buf, engine='openpyxl') as writer:
                            master_df_with_total.to_excel(writer, index=False, sheet_name='Master_Cadre_Status')
                        st.download_button(
                            "📥 Download Full District Cadre-wise Master Excel", 
                            data=master_buf.getvalue(), 
                            file_name="All_Mandals_Cadre_Wise_Status.xlsx",
                            use_container_width=True
                        )
                    with col_all2:
                        render_print_button(master_df_with_total.head(200), report_title="FULL DISTRICT CADRE MASTER REPORT", subtitle="All Mandals - Cadre Status")
            else:
                st.info("Cadre vivaraalu labhinchaledhu.")

# ==========================================
# ------------ TAB 4: MIS REPORTS ---------
# ==========================================
with tab4:
    st.info("📄 CSE MIS Reports - Module Coming Soon")
