import streamlit as st
import pandas as pd
from datetime import date, timedelta
from docx import Document
import base64
import os

@st.cache_data
def load_data():
    master_df = pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name=None)
    sf11_df = pd.read_excel("assets/SF-11 Register.xlsm", sheet_name=None)
    return master_df, sf11_df

def replace_placeholders(doc, context):
    for p in doc.paragraphs:
        for key, val in context.items():
            if f"[{key}]" in p.text:
                p.text = p.text.replace(f"[{key}]", str(val))
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, val in context.items():
                    if f"[{key}]" in cell.text:
                        cell.text = cell.text.replace(f"[{key}]", str(val))

def generate_docx(template_path, context, filename):
    doc = Document(template_path)
    replace_placeholders(doc, context)
    docx_path = os.path.join("/tmp", filename + ".docx")
    doc.save(docx_path)
    return docx_path

def convert_to_pdf(docx_path):
    try:
        from docx2pdf import convert
        pdf_path = docx_path.replace(".docx", ".pdf")
        convert(docx_path, pdf_path)
        return pdf_path
    except:
        return None

def download_button(file_path, label):
    with open(file_path, "rb") as f:
        data = f.read()
        b64 = base64.b64encode(data).decode()
        href = f'<a href="data:application/octet-stream;base64,{b64}" download="{os.path.basename(file_path)}">{label}</a>'
        st.markdown(href, unsafe_allow_html=True)

def correct_unit_station_format(unit_val, station_val):
    try:
        unit_str = str(unit_val).strip().split("/")[0]
        if unit_str.isdigit():
            unit_str = unit_str[:2]
        station_str = str(station_val).strip()
        return f"{unit_str}/{station_str}"
    except:
        return ""

# Load data
master_data, sf11_data = load_data()
master_sheet = list(master_data.keys())[0]
sf11_sheet = list(sf11_data.keys())[0]
df_master = master_data[master_sheet]
df_sf11 = sf11_data[sf11_sheet]

col_pf = 1
col_hrms = 2
col_unit = 4
col_english_name = 5
col_hindi_name = 13
col_designation = 18
col_station = 9

col_sf11_name = 2
col_sf11_letter_no = 5
col_sf11_memo = 6

# Letter type
letter_type = st.selectbox("Select Letter Type:", [
    "SF-11 Punishment Order",
    "Duty Letter (For Absent)",
    "Sick Memo",
    "Exam NOC"
])

# Select employee dropdown depending on letter type
if letter_type == "SF-11 Punishment Order":
    df_sf11.dropna(subset=[df_sf11.columns[col_sf11_name], df_sf11.columns[col_sf11_letter_no]], inplace=True)
    df_sf11["Dropdown"] = df_sf11.apply(lambda row: f"{row[col_sf11_name]} / {row[col_sf11_letter_no]}", axis=1)
    emp_options = df_sf11["Dropdown"].tolist()
    selected_emp = st.selectbox("Select Employee from SF-11 Register:", emp_options)
    row = df_sf11[df_sf11["Dropdown"] == selected_emp].iloc[0]
    hindi_name = row[col_sf11_name]
    memo_text = row[col_sf11_memo]
    designation = ""
    unit_number = ""
    pf_number = ""
else:
    df_master["DisplayUnit"] = df_master.apply(lambda row: correct_unit_station_format(row[col_unit], row[col_station]), axis=1)
    df_master["Dropdown"] = df_master.apply(lambda row: f"{row[col_pf]} - {row[col_hrms]} - {row['DisplayUnit']} - {row[col_english_name]}", axis=1)
    emp_options = df_master["Dropdown"].tolist()
    selected_emp = st.selectbox("Select Employee from Master Data:", emp_options)
    row = df_master[df_master["Dropdown"] == selected_emp].iloc[0]
    hindi_name = row[col_hindi_name]
    memo_text = st.text_area("Memo Text") if letter_type == "SF-11 Punishment Order" else ""
    designation = row[col_designation]
    unit_number = row["DisplayUnit"]
    pf_number = row[col_pf]

letter_date = st.date_input("Select Letter Date", date.today())
from_date = st.date_input("From Date") if "Duty" in letter_type else None
to_date = st.date_input("To Date") if "Duty" in letter_type else None
duty_date_default = (to_date + timedelta(days=1)) if to_date else date.today()
duty_date = st.date_input("Join Duty Date", duty_date_default) if "Duty" in letter_type else None
exam_name = st.text_input("Exam Name") if "NOC" in letter_type else ""
noc_count = st.selectbox("NOC Attempt No", [1, 2, 3, 4]) if "NOC" in letter_type else None

context = {
    "LetterDate": letter_date.strftime("%d-%m-%Y"),
    "EmployeeName": hindi_name,
    "Designation": designation,
    "UnitNumber": unit_number,
    "FromDate": from_date.strftime("%d-%m-%Y") if from_date else "",
    "ToDate": to_date.strftime("%d-%m-%Y") if to_date else "",
    "DutyDate": duty_date.strftime("%d-%m-%Y") if duty_date else "",
    "MEMO": memo_text,
    "PFNumber": pf_number,
    "ExamName": exam_name,
    "NOCCount": noc_count
}

template_files = {
    "SF-11 Punishment Order": "assets/SF-11 Punishment order temp.docx",
    "Duty Letter (For Absent)": "assets/Absent Duty letter temp.docx",
    "Sick Memo": "assets/SICK MEMO temp..docx",
    "Exam NOC": "assets/Exam NOC Letter temp.docx"
}

if st.button("Generate Letter"):
    base_filename = f"{letter_type.split()[0]}_{hindi_name}_{letter_date.strftime('%d-%m-%Y')}"
    docx_path = generate_docx(template_files[letter_type], context, base_filename)
    st.success("Word letter generated successfully.")
    download_button(docx_path, f"⬇️ Download {os.path.basename(docx_path)}")

    pdf_path = convert_to_pdf(docx_path)
    if pdf_path and os.path.exists(pdf_path):
        st.success("PDF letter generated successfully.")
        download_button(pdf_path, f"⬇️ Download {os.path.basename(pdf_path)}")
    else:
        st.warning("PDF conversion not supported on this platform.")
