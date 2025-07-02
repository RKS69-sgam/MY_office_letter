import streamlit as st
import pandas as pd
from datetime import date, timedelta
from docx import Document
from tempfile import NamedTemporaryFile
import base64
import os

# ====== Load Employee Data ======
@st.cache_data
def load_employee_data():
    df = pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name=None)
    return df

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

def generate_docx(template_path, context):
    doc = Document(template_path)
    replace_placeholders(doc, context)
    temp_file = NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(temp_file.name)
    return temp_file.name

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

# ========== Column Mappings ==========
col_pf = 1
col_hrms = 2
col_unit_raw = 4
col_emp_eng = 5
col_station = 8
col_emp_hin = 13
col_designation = 18  # Col 19 (0-based)
# =====================================

data = load_employee_data()
sheet_names = list(data.keys())
selected_sheet = st.selectbox("Select Unit Sheet:", sheet_names)
df = data[selected_sheet]

# Drop-down list with Employee Name (Eng), PF, HRMS, Unit
def format_label(row):
    return f"{row[col_emp_eng]} (PF: {row[col_pf]}, HRMS: {row[col_hrms]}, Unit: {row[col_unit_raw]})"

employee_labels = df.apply(format_label, axis=1).tolist()
selected_label = st.selectbox("Select Employee (with details):", employee_labels)
selected_index = employee_labels.index(selected_label)
selected_row = df.iloc[selected_index]

# Letter type
letter_type = st.selectbox("Select Letter Type:", [
    "SF-11 Punishment Order",
    "Duty Letter (For Absent)",
    "Sick Memo",
    "Exam NOC"
])

letter_date = st.date_input("Select Letter Date", date.today())
from_date = st.date_input("From Date") if "Duty" in letter_type else None
to_date = st.date_input("To Date") if "Duty" in letter_type else None

# Auto Join Date = ToDate + 1 for Absent Duty
if "Duty" in letter_type:
    default_join = to_date + timedelta(days=1) if to_date else date.today()
    duty_date = st.date_input("Join Duty Date", default_join)
else:
    duty_date = None

memo_text = st.text_area("Memo Text") if "SF-11" in letter_type else ""
exam_name = st.text_input("Exam Name") if "NOC" in letter_type else ""
noc_count = st.selectbox("NOC Attempt No", [1, 2, 3, 4]) if "NOC" in letter_type else None

# ===== Process Unit Number with Station =====
raw_unit = str(selected_row[col_unit_raw])
working_station = str(selected_row[col_station])
if "/" in raw_unit:
    unit_part = raw_unit.split("/")[0]
elif raw_unit.isdigit():
    unit_part = raw_unit[:2]
else:
    unit_part = raw_unit
unit_combined = f"{unit_part}/{working_station}"

# ===== Prepare context =====
context = {
    "LetterDate": letter_date.strftime("%d-%m-%Y"),
    "EmployeeName": selected_row[col_emp_hin],
    "Designation": selected_row[col_designation] if len(selected_row) > col_designation else "",
    "UnitNumber": unit_combined,
    "FromDate": from_date.strftime("%d-%m-%Y") if from_date else "",
    "ToDate": to_date.strftime("%d-%m-%Y") if to_date else "",
    "DutyDate": duty_date.strftime("%d-%m-%Y") if duty_date else "",
    "MEMO": memo_text,
    "PFNumber": selected_row[col_pf],
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
    docx_path = generate_docx(template_files[letter_type], context)
    st.success("Word letter generated successfully.")
    download_button(docx_path, "⬇️ Download Word Letter")

    pdf_path = convert_to_pdf(docx_path)
    if pdf_path and os.path.exists(pdf_path):
        st.success("PDF letter generated successfully.")
        download_button(pdf_path, "⬇️ Download PDF Letter")
    else:
        st.warning("PDF conversion not supported on this platform.")