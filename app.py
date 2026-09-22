import streamlit as st
import pandas as pd
import os
import io

st.set_page_config(page_title="District Education Portal", layout="wide")

st.title("🎓 District Educational Office - Management Portal")
st.caption("West Godavari District - School Education Department")
st.markdown("---")

EXCEL_FILE_PATH = "UPTO DATE UDISE ROLL.xlsx"

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
        st.error(f"Error reading Excel file: {e}")
        return None

df = load_udise_data(EXCEL_FILE_PATH)

tab1, tab2, tab3, tab4 = st.tabs([
    "🏫 School 360° & UDISE Reports",
    "🧑‍🏫 Teachers Directory & Retirement",
    "📊 Cadre Strength & Vacancy",
    "📄 CSE MIS Reports"
])

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
                    
                    # Class-wise Enrolment Breakdown Section
                    st.markdown("#### 📋 Class-wise Enrolment Breakdown")
                    
                    # Collect all class related columns
                    class_data = []
                    # Common class prefixes: PP3, PP2, PP1, C1 to C12 or I to XII
                    all_cols = list(df.columns)
                    
                    # Identify distinct class keys from headers like PP3(Boys), Class 1(Boys), etc.
                    class_keys = []
                    for c in all_cols:
                        if '(' in c and ')' in c:
                            key = c.split('(')[0].strip()
                            if key not in class_keys and any(tag in c.upper() for tag in ['BOY', 'GIRL', 'TOTAL', 'TRANS']):
                                class_keys.append(key)

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
                        st.info("ఈ పాఠశాలకు సంబంధించిన తరగతుల వారీ వివరాలు షీట్‌లో గుర్తించబడలేదు.")
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
                st.download_button(
                    label="📥 Download Mandal Abstract as Excel",
                    data=m_buf.getvalue(),
                    file_name="Mandal_Wise_Abstract.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Block Name లేదా Total Enrolment కాలమ్ గుర్తించబడలేదు.")

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
            st.download_button(
                label="📥 Download Filtered Report as Excel",
                data=buffer.getvalue(),
                file_name="Filtered_UDISE_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

with tab2:
    st.info("🧑‍🏫 Teachers Directory & Retirement Tracker - Module Coming Soon")

with tab3:
    st.info("📊 Cadre Strength & Vacancy Analysis - Module Coming Soon")

with tab4:
    st.info("📄 CSE MIS Reports - Module Coming Soon")
