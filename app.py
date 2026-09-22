import streamlit as st
import pandas as pd
import os
import io
import shutil

st.set_page_config(page_title="District Education Portal", layout="wide")

st.title("🎓 District Educational Office - Management Portal")
st.caption("West Godavari District - School Education Department")
st.markdown("---")

EXCEL_FILE_PATH = "C:\Users\admin\OneDrive - AP School Education\EDUCATION DATA\UPTO DATE UDISE ROLL.xlsx"
tab1, tab2, tab3, tab4 = st.tabs([
    "🏫 School 360° & UDISE Reports", 
    "👨‍🏫 Teachers Directory & Retirement", 
    "📊 Cadre Strength & Vacancy", 
    "📑 CSE MIS Reports"
])

@st.cache_data(ttl=30)
def load_udise_data(file_path):
    if not os.path.exists(file_path):
        return None, f"File dorakaledhu: {file_path}"
    try:
        temp_dir = os.path.join(os.environ.get('TEMP', '.'), "deo_portal_cache")
        os.makedirs(temp_dir, exist_ok=True)
        temp_copy = os.path.join(temp_dir, "roll_data.xlsx")
        shutil.copyfile(file_path, temp_copy)
        df = pd.read_excel(temp_copy)
        return df, None
    except Exception:
        try:
            with open(file_path, 'rb') as f:
                content = f.read()
            df = pd.read_excel(io.BytesIO(content))
            return df, None
        except Exception as e:
            return None, str(e)

# Helper function to convert dataframe to Excel file in memory
def to_excel(df_to_export):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_to_export.to_excel(writer, index=False, sheet_name='Report')
    processed_data = output.getvalue()
    return processed_data

# ----------------- TAB 1: School Profile & Reports -----------------
with tab1:
    df_roll, err = load_udise_data(EXCEL_FILE_PATH)
    
    if err:
        st.error(f"⚠️ {err}")
    elif df_roll is not None:
        df_roll['UDISE Code'] = df_roll['UDISE Code'].astype(str).str.strip().str.replace(".0", "", regex=False)
        
        sub_tab1, sub_tab2 = st.tabs(["🔍 School 360° Search", "📑 Custom Reports & Excel Export"])
        
        # --- SUB TAB 1: Single School Search ---
        with sub_tab1:
            st.subheader("🏫 Individual School 360° Profile")
            col_search, col_btn = st.columns([3, 1])
            with col_search:
                search_code = st.text_input("Enter 11 Digit UDISE Code:", placeholder="e.g., 28153500204", key="sch_search")
            with col_btn:
                st.write("")
                st.write("")
                search_btn = st.button("Search Profile", use_container_width=True)
                
            if search_code:
                clean_search = str(search_code).strip()
                row = df_roll[df_roll['UDISE Code'] == clean_search]
                
                if not row.empty:
                    s = row.iloc[0]
                    st.markdown(f"""
                    <div style="background-color: #1E293B; padding: 20px; border-radius: 10px; border-left: 6px solid #2563EB; margin-bottom: 20px;">
                        <h2 style="color: #FFFFFF; margin: 0 0 10px 0;">🏫 {s.get('School Name', 'N/A')}</h2>
                        <div style="color: #CBD5E1; font-size: 15px; display: flex; flex-wrap: wrap; gap: 20px;">
                            <span><b>UDISE:</b> {s.get('UDISE Code', 'N/A')}</span>
                            <span><b>Mandal:</b> {s.get('Block Name', 'N/A')}</span>
                            <span><b>Cluster:</b> {s.get('Cluster Code', 'N/A')}</span>
                            <span><b>Management:</b> {s.get('School Management', 'N/A')}</span>
                            <span><b>Category:</b> {s.get('School Category', 'N/A')}</span>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    m1, m2, m3, m4 = st.columns(4)
                    gt_val = int(pd.to_numeric(s.get('Grand Total', 0), errors='coerce') or 0)
                    b_val = int(pd.to_numeric(s.get('Total Boys', 0), errors='coerce') or 0)
                    g_val = int(pd.to_numeric(s.get('Total Girls', 0), errors='coerce') or 0)
                    t_val = int(pd.to_numeric(s.get('Total Trans', 0), errors='coerce') or 0)
                    
                    m1.metric("Grand Total Enrolment", f"{gt_val:,}")
                    m2.metric("Total Boys 👦", f"{b_val:,}")
                    m3.metric("Total Girls 👧", f"{g_val:,}")
                    m4.metric("Transgender", f"{t_val:,}")
                    
                    st.markdown("---")
                    class_names = [
                        "PP3", "PP2", "PP1", "Class 1", "Class 2", "Class 3", "Class 4", "Class 5",
                        "Class 6", "Class 7", "Class 8", "Class 9", "Class 10", "Class 11", "Class 12"
                    ]
                    class_data = []
                    for c in class_names:
                        tot_col = f"{c}(Total)"
                        if tot_col in s:
                            class_data.append({
                                "Class": c,
                                "Boys": int(pd.to_numeric(s.get(f"{c}(Boys)", 0), errors='coerce') or 0),
                                "Girls": int(pd.to_numeric(s.get(f"{c}(Girls)", 0), errors='coerce') or 0),
                                "Transgender": int(pd.to_numeric(s.get(f"{c}(Trans)", 0), errors='coerce') or 0),
                                "Total Strength": int(pd.to_numeric(s.get(tot_col, 0), errors='coerce') or 0)
                            })
                    df_cls = pd.DataFrame(class_data)
                    act_cls = df_cls[df_cls["Total Strength"] > 0].reset_index(drop=True)
                    
                    c_left, c_right = st.columns([3, 2])
                    with c_left:
                        st.markdown("##### 📋 Class-wise Enrolment Breakdown")
                        st.dataframe(act_cls if not act_cls.empty else df_cls, use_container_width=True, hide_index=True)
                    with c_right:
                        st.markdown("##### 📊 Boys vs Girls Comparison")
                        if not act_cls.empty:
                            st.bar_chart(act_cls.set_index("Class")[["Boys", "Girls"]])
                else:
                    st.warning(f"⚠️ UDISE Code `{clean_search}` tho school kanipinchaledhu.")

        # --- SUB TAB 2: Custom Reports & Excel Export ---
        with sub_tab2:
            st.subheader("📑 Custom Report Generator & Excel Export")
            
            f_col1, f_col2, f_col3 = st.columns(3)
            
            # Mandal Filter
            all_mandals = ["All Mandals"] + sorted(list(df_roll['Block Name'].dropna().unique()))
            with f_col1:
                sel_mandal = st.selectbox("1. Filter by Mandal (Block Name):", all_mandals)
            
            # Management Filter
            all_mgmt = ["All Managements"] + sorted(list(df_roll['School Management'].dropna().unique()))
            with f_col2:
                sel_mgmt = st.selectbox("2. Filter by Management:", all_mgmt)
                
            # Roll Range Filter
            with f_col3:
                roll_filter = st.selectbox("3. Enrolment Range:", [
                    "All Schools", 
                    "Low Roll (< 30 Students)", 
                    "Medium Roll (30 to 100 Students)", 
                    "High Roll (> 100 Students)"
                ])
                
            # Filtering logic
            filtered_df = df_roll.copy()
            if sel_mandal != "All Mandals":
                filtered_df = filtered_df[filtered_df['Block Name'] == sel_mandal]
            if sel_mgmt != "All Managements":
                filtered_df = filtered_df[filtered_df['School Management'] == sel_mgmt]
                
            filtered_df['Grand Total'] = pd.to_numeric(filtered_df['Grand Total'], errors='coerce').fillna(0).astype(int)
            if roll_filter == "Low Roll (< 30 Students)":
                filtered_df = filtered_df[filtered_df['Grand Total'] < 30]
            elif roll_filter == "Medium Roll (30 to 100 Students)":
                filtered_df = filtered_df[(filtered_df['Grand Total'] >= 30) & (filtered_df['Grand Total'] <= 100)]
            elif roll_filter == "High Roll (> 100 Students)":
                filtered_df = filtered_df[filtered_df['Grand Total'] > 100]
                
            # Show Filter Results
            st.markdown("---")
            rf1, rf2, rf3 = st.columns(3)
            rf1.metric("Total Filtered Schools", len(filtered_df))
            rf2.metric("Total Students in List", f"{filtered_df['Grand Total'].sum():,}")
            
            # Download Button
            with rf3:
                st.write("")
                excel_bytes = to_excel(filtered_df)
                st.download_button(
                    label="📥 Download This Report (Excel)",
                    data=excel_bytes,
                    file_name=f"UDISE_Report_{sel_mandal}_{roll_filter}.xlsx".replace(" ", "_"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            # Summary Table view
            display_columns = [
                'UDISE Code', 'School Name', 'Block Name', 'School Management', 
                'School Category', 'Total Boys', 'Total Girls', 'Grand Total'
            ]
            valid_cols = [c for c in display_columns if c in filtered_df.columns]
            st.dataframe(filtered_df[valid_cols], use_container_width=True, hide_index=True)

# ----------------- TAB 2: Teachers Data -----------------
with tab2:
    st.subheader("👨‍🏫 Teachers Directory & Retirement Tracker")
    st.info("Teachers details file thvaralo link cheddam.")

# ----------------- TAB 3: Cadre Strength -----------------
with tab3:
    st.subheader("📊 Cadre Strength vs Working vs Vacancy")
    st.info("Cadre strength report module.")

# ----------------- TAB 4: CSE MIS Reports -----------------
with tab4:
    st.subheader("📑 CSE MIS Custom Formatter")
    st.info("MIS reports formatting module.")
