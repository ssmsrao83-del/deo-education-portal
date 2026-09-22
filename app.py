import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
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
          var printWin = window.open('', '', 'width=950,height=700');
          printWin.document.write(`
            <html>
            <head>
              <title>{report_title}</title>
              <style>
                body {{ font-family: Arial, sans-serif; margin: 20px; color: #000; }}
                .header-area {{ text-align: center; border-bottom: 2px solid #000; padding-bottom: 8px; margin-bottom: 15px; }}
                .header-area h2 {{ margin: 0; font-size: 17px; color: #1a237e; text-transform: uppercase; }}
                .header-area h3 {{ margin: 4px 0; font-size: 14px; color: #333; }}
                .header-area p {{ margin: 3px 0 0 0; font-size: 12px; color: #555; font-weight: bold; }}
                .print-table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
                .print-table th, .print-table td {{ border: 1px solid #444; padding: 6px 8px; text-align: left; }}
                .print-table th {{ background-color: #f2f2f2 !important; font-weight: bold; -webkit-print-color-adjust: exact; }}
                .print-table tr:nth-child(even) {{ background-color: #fafafa; -webkit-print-color-adjust: exact; }}
                .print-table tr:last-child {{ font-weight: bold; background-color: #eaeaea !important; border-top: 2px solid #222; }}
                @media print {{ body {{ margin: 10mm; }} .print-table {{ font-size: 11px; }} }}
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
            if 'CATEG' in col.upper():
                df['Category_Display'] = df[col].apply(lambda x: clean_and_map(x, CATEGORY_MAPPING))
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
        df_m = pd.read_excel(actual_path)
        # Normalize column names: remove hidden unicode characters and strip extra spaces
        df_m.columns = [" ".join(str(c).replace('\xa0', ' ').split()).strip() for c in df_m.columns]
        
        # Management and Category display mapping
        for col in df_m.columns:
            if 'MANAGEMENT' in col.upper():
                df_m['Management_Display'] = df_m[col].apply(lambda x: clean_and_map(x, MANAGEMENT_MAPPING))
            if 'CATEGORY' in col.upper():
                df_m['Category_Display'] = df_m[col].apply(lambda x: clean_and_map(x, CATEGORY_MAPPING))

        # Explicit target matching for 5-15 and 15 and above
        p5_15_col = next((c for c in df_m.columns if '5-15' in c), None)
        p15_p_col = next((c for c in df_m.columns if '15 AND ABOVE' in c.upper() or '15 AND' in c.upper()), None)
        
        if p5_15_col:
            df_m[p5_15_col] = pd.to_numeric(df_m[p5_15_col], errors='coerce').fillna(0)
        if p15_p_col:These headers represent standard data validation fields used in school education reporting and the UDISE+ (Unified District Information System for Education Plus) portal for tracking student identity verification and Mandatory Biometric Updates (MBU). 

Here is a breakdown of what each column tracks:

### Administrative & School Identification
* **District Name / District Code:** Administrative district of the institution.
* **Block Name / Block Code:** Educational block/mandal level administrative unit.
* **School Name / UDISE Code:** School title and its unique 11-digit national identification code.
* **School Management:** Governing authority (e.g., Department of Education, Local Body, Private Unaided, Aided, Central Government).
* **School Category:** Level of education offered (e.g., Primary, Upper Primary, Secondary, Higher Secondary).
* **Academic Year:** Reporting session (e.g., 2025–2026).

---

### Student Enrollment & Demographic Verification
* **Total Student:** Total active enrollment recorded in the school for that session.
* **AADHAAR Provided:** Number of students whose ID details have been seeded into the system.
* **AADHAAR Verified-Passed (Yes):** Records where the demographic details (Name, Gender, Date of Birth) exactly match the central database.
* **AADHAAR Verified-Failed (No):** Records where verification failed due to demographic mismatches (spelling differences, DOB discrepancy, gender error).
* **AADHAAR Verified-To be Done:** Records seeded or uploaded where automated verification has not yet run.

---

### Mandatory Biometric Update (MBU) Tracking
UIDAI mandates updating child biometrics (fingerprints, iris, and facial photo) at age 5 and age 15:
* **MBU Pending (Age 5–15):** Students between ages 5 and 15 who have reached age 5 but have not completed their first mandatory biometric update.
* **MBU Pending (Age 15 and above):** Students aged 15 or older who have not completed their final mandatory biometric update.
* **MBU Not Required:** Students whose biometric updates are fully compliant and current for their age group.
* **MBU Not Applicable:** Special categories, exemptions, or student records where MBU tracking does not apply.
* **Status Check to be done:** Records flagged for manual audit, reconciliation, or pending API response checks between state systems and the central database.

---

### Recommended Action Items for Resolution

1. **For "Verified-Failed":** Download the school error log, compare the student admission register against the physical ID card, and correct the Name/DOB/Gender spellings in the portal before re-submitting for automated verification.
2. **For "MBU Pending":** Direct students/parents to designated enrolment centers, permanent update centers, or school-level special camps to complete the biometric capture.
3. **For "To be Done":** Trigger the bulk verification sync on the portal dashboard.
