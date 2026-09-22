import streamlit as st
import pandas as pd
import os
import io

st.set_page_config(page_title="District Education Portal", layout="wide")

st.title("🎓 District Educational Office - Management Portal")
st.caption("West Godavari District - School Education Department")
st.markdown("---")

EXCEL_FILE_PATH = "UPTO DATE UDISE ROLL.xlsx"
TEACHERS_FILE_PATH = "TEACHERS DATA.xlsx"

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

def clean_and_map(val, mapping_dict):
    try:
        val_str = str(val).split('-')[0].strip()
        num = int(float(val_str))
        return mapping_dict.get(num, str(val))
    except:
        return str(val)

@st.cache_data(ttl=30)
def load_udise_data(file_path):
    if not os.path.exists(file_path):
        return None
    try:
        df = pd.read_excel(file_path)
        df.columns = [str(c).strip() for c in df.columns]
        for col in df.columns:
            if 'MANAGE' in col.upper():
                df['Management_Display'] = df[col].apply(lambda x: clean_and_map(x, MANAGEMENT_MAPPING))
            if 'CATEG' in col.upper():
                df['Category_Display'] = df[col].apply(lambda x: clean_and_map(x, CATEGORY_MAPPING))
        return df
    except Exception as e:
        st.error(f"Error reading UDISE file: {e}")
        return None

@st.cache_data(ttl=30)
def load_cadre_data(file_path):
    if not os.path.exists(file_path):
        return None
    try:
        df_raw = pd.read_excel(file_path, header=[0, 1])
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
        # De-duplicate any column names if needed
        df_raw = df_raw.loc[:, ~df_raw.columns.duplicated()]
        return df_raw
    except Exception as e:
        try:
            df_fallback = pd.read_excel(file_path)
            df_fallback = df_fallback.loc[:, ~df_fallback.columns.duplicated()]
            return df_fallback
        except:
            return None

df = load_udise_data(EXCEL_FILE_PATH)
df_cadre = load_cadre_data(TEACHERS_FILE_PATH)

tab1, tab2, tab3, tab4 = st.tabs([
    "🏫 School 360° & UDISE Reports",
    "🧑‍🏫 Teachers Directory & Retirement",
    "📊 Cadre Strength & Vacancy",
    "📄 CSE MIS Reports"
])

# ----------------- TAB 1: UDISE -----------------
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

        subtab1, subtab2, subtab3 = st.tabs([
            "🔍 School 360° Search", 
            "📊 Mandal-wise Abstract", 
            "📑 Custom Reports & Excel Export"
        ])

        with subtab1:
            st.subheader("🏫 Individual School 360° Profile")
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
                        col_t1, col_t2 = st.columns([3, 2])
                        with col_t1:
                            st.dataframe(cdf, use_container_width=True, hide_index=True)
                        with col_t2:
                            st.bar_chart(cdf.set_index("Class")[["Boys 👦", "Girls 👧"]])
                    else:
                        st.info("ఈ పాఠశాలకు సంబంధించిన తరగతుల వివరాలు లభించలేదు.")
                else:
                    st.warning("ఈ UDISE కోడ్ తో రికార్డు కనబడలేదు.")

        with subtab2:
            st.subheader("📊 Mandal-wise Enrolment Abstract")
            if block_col and tot_col:
                agg_dict = {udise_col: 'count', tot_col: 'sum'}
                rename_cols = {udise_col: 'Total Schools', tot_col: 'Total Enrolment'}
                if boys_col:
                    agg_dict[boys_col] = 'sum'
                    rename_cols[boys_col] = 'Total Boys'
                if girls_col:
                    agg_dict[girls_col] = 'sum'
                    rename_cols[girls_col] = 'Total Girls'
                
                mandal_summary = df.groupby(block_col).agg(agg_dict).reset_index()
                mandal_summary = mandal_summary.rename(columns={block_col: 'Mandal (Block)'})
                mandal_summary = mandal_summary.rename(columns=rename_cols)
                st.dataframe(mandal_summary, use_container_width=True)
                
                m_buf = io.BytesIO()
                with pd.ExcelWriter(m_buf, engine='openpyxl') as writer:
                    mandal_summary.to_excel(writer, index=False, sheet_name='Mandal_Abstract')
                st.download_button("📥 Download Mandal Abstract as Excel", data=m_buf.getvalue(), file_name="Mandal_Wise_Abstract.xlsx")

        with subtab3:
            st.subheader("📑 Custom Reports & Excel Export")
            f1, f2 = st.columns(2)
            with f1:
                mandal_list = sorted(list(df[block_col].dropna().unique())) if block_col else []
                sel_mandals = st.multiselect("Select Mandal(s):", mandal_list, default=mandal_list)
            with f2:
                mgmt_list = sorted(list(df[mgmt_col].dropna().unique())) if mgmt_col else []
                sel_mgmt = st.multiselect("Select Management:", mgmt_list, default=mgmt_list)

            filtered_df = df.copy()
            if block_col and sel_mandals:
                filtered_df = filtered_df[filtered_df[block_col].isin(sel_mandals)]
            if mgmt_col and sel_mgmt:
                filtered_df = filtered_df[filtered_df[mgmt_col].isin(sel_mgmt)]

            st.write(f"మొత్తం పాఠశాలలు: **{len(filtered_df)}**")
            st.dataframe(filtered_df, use_container_width=True)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                filtered_df.to_excel(writer, index=False, sheet_name='Filtered_Report')
            st.download_button("📥 Download Filtered Report as Excel", data=buffer.getvalue(), file_name="Filtered_UDISE_Report.xlsx")

# ----------------- TAB 2: TEACHERS -----------------
with tab2:
    st.info("🧑‍🏫 Teachers Directory & Retirement Tracker - Module Coming Soon")

# ----------------- TAB 3: CADRE STRENGTH & VACANCY -----------------
with tab3:
    st.subheader("📊 Cadre Strength, Working & Vacancy Analysis")
    if df_cadre is None:
        st.warning(f"⚠️ '{TEACHERS_FILE_PATH}' ఫైల్ సరిగ్గా లోడ్ కాలేదు. దయచేసి ఫైల్ అప్‌లోడ్ అయిందో లేదో తనిఖీ చేయండి.")
    else:
        # Robust column search
        c_udise = next((c for c in df_cadre.columns if 'UDISE' in str(c).upper()), None)
        c_school = next((c for c in df_cadre.columns if any(k in str(c).upper() for k in ['HS/UPS_NAME', 'SCHOOL', 'NAME'])), None)
        c_mandal = next((c for c in df_cadre.columns if 'MANDAL' in str(c).upper()), None)
        
        # Categorize columns into Sanctioned, Working, Vacant
        sanc_cols = [c for c in df_cadre.columns if 'SANCTIONED' in str(c).upper() and 'TOTAL' not in str(c).upper()]
        work_cols = [c for c in df_cadre.columns if 'WORKING' in str(c).upper() and 'TOTAL' not in str(c).upper() and 'MTS' not in str(c).upper()]
        vac_cols  = [c for c in df_cadre.columns if 'VACANT' in str(c).upper() and 'TOTAL' not in str(c).upper()]
        
        tot_sanc_col = next((c for c in df_cadre.columns if 'SANCTIONED' in str(c).upper() and 'TOTAL' in str(c).upper()), None)
        tot_work_col = next((c for c in df_cadre.columns if 'WORKING' in str(c).upper() and 'TOTAL' in str(c).upper()), None)
        tot_vac_col  = next((c for c in df_cadre.columns if 'VACANT' in str(c).upper() and 'TOTAL' in str(c).upper()), None)

        c_tab1, c_tab2, c_tab3 = st.tabs([
            "🔍 School Cadre Profile", 
            "📌 Subject/Post Vacancies", 
            "📑 Mandal Cadre Summary"
        ])

        with c_tab1:
            st.markdown("#### 🏫 Individual School Cadre Strength")
            cadre_search = st.text_input("Enter UDISE Code or School Name:", value="28153500204", key="cadre_srch_input")
            
            matched_cadre = pd.DataFrame()
            if cadre_search and c_udise:
                # Safe single series extraction
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
                for sc in sanc_cols:
                    post_name = sc.split('-')[-1].strip() if '-' in sc else sc
                    wc = next((c for c in work_cols if post_name.upper() in c.upper()), None)
                    vc = next((c for c in vac_cols if post_name.upper() in c.upper()), None)
                    
                    s_val = int(pd.to_numeric(c_row[sc], errors='coerce')) if pd.notnull(c_row[sc]) else 0
                    w_val = int(pd.to_numeric(c_row[wc], errors='coerce')) if wc and pd.notnull(c_row[wc]) else 0
                    vac_val = int(pd.to_numeric(c_row[vc], errors='coerce')) if vc and pd.notnull(c_row[vc]) else 0
                    
                    if s_val > 0 or w_val > 0 or vac_val > 0:
                        post_list.append({
                            "Designation / Cadre": post_name,
                            "Sanctioned": s_val,
                            "Working": w_val,
                            "Vacant": vac_val
                        })
                
                if post_list:
                    p_df = pd.DataFrame(post_list)
                    st.dataframe(p_df, use_container_width=True, hide_index=True)
                else:
                    st.info("ఈ పాఠశాలకు సంబంధించిన పోస్టుల విభజన వివరాలు అందుబాటులో లేవు.")
            else:
                st.info("పాఠశాల వివరాలు చూడటానికి UDISE కోడ్ నమోదు చేయండి.")

        with c_tab2:
            st.markdown("#### 📌 District / Mandal-wise Vacancy by Subject & Cadre")
            if vac_cols:
                vac_summary = []
                for vc in vac_cols:
                    p_label = vc.split('-')[-1].strip() if '-' in vc else vc
                    total_v = int(pd.to_numeric(df_cadre[vc], errors='coerce').fillna(0).sum())
                    if total_v > 0:
                        vac_summary.append({"Designation / Post": p_label, "Total Vacancies": total_v})
                
                if vac_summary:
                    vdf = pd.DataFrame(vac_summary).sort_values(by="Total Vacancies", ascending=False)
                    v_col1, v_col2 = st.columns([2, 3])
                    with v_col1:
                        st.dataframe(vdf, use_container_width=True, hide_index=True)
                    with v_col2:
                        st.bar_chart(vdf.set_index("Designation / Post"))
                else:
                    st.info("ఎటువంటి ఖాళీలు నమోదు కాలేదు.")

        with c_tab3:
            st.markdown("#### 📑 Mandal-wise Cadre Abstract & Download")
            if c_mandal and tot_sanc_col and tot_work_col and tot_vac_col:
                for c in [tot_sanc_col, tot_work_col, tot_vac_col]:
                    df_cadre[c] = pd.to_numeric(df_cadre[c], errors='coerce').fillna(0)
                
                m_cadre = df_cadre.groupby(c_mandal).agg({
                    c_udise: 'count',
                    tot_sanc_col: 'sum',
                    tot_work_col: 'sum',
                    tot_vac_col: 'sum'
                }).reset_index()
                
                m_cadre.columns = ['Mandal', 'Schools', 'Total Sanctioned', 'Total Working', 'Total Vacant']
                st.dataframe(m_cadre, use_container_width=True)
                
                cadre_buf = io.BytesIO()
                with pd.ExcelWriter(cadre_buf, engine='openpyxl') as writer:
                    m_cadre.to_excel(writer, index=False, sheet_name='Cadre_Abstract')
                st.download_button(
                    "📥 Download Cadre Abstract Excel", 
                    data=cadre_buf.getvalue(), 
                    file_name="Mandal_Cadre_Abstract.xlsx"
                )

# ----------------- TAB 4: MIS REPORTS -----------------
with tab4:
    st.info("📄 CSE MIS Reports - Module Coming Soon")
