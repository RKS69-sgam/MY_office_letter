import streamlit as st
import pandas as pd
from datetime import date
from docx import Document
from io import BytesIO
from tempfile import NamedTemporaryFile
import base64
import os

# Load Employee Data
@st.cache_data
def load_employee_data():
    df = pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name=None)
    return df

def generate_docx(template_path, context):
    doc = Document(template_path)
    for p in doc.paragraphs:
        for key, val in context.items():
            if f"[{key}]" in p.text:
                p.text = p.text.replace(f"[{key}]", str(val))
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
        href = f'<a href="data:application/octet-stream;base64,{b64}" download="{file_path.split("/")[-1]}">{label}</a>'
        st.markdown(href, unsafe_allow_html=True)

data = load_employee_data()
sheet_names = list(data.keys())
selected_sheet = st.selectbox("Select Unit Sheet:", sheet_names)
df = data[selected_sheet]

# Employee Selection
employee_names = df.iloc[:, 4].dropna().tolist()
selected_emp = st.selectbox("Select Employee:", employee_names)
selected_row = df[df.iloc[:, 4] == selected_emp].iloc[0]

# Letter Type
letter_type = st.selectbox("Select Letter Type:", [
    "SF-11 Punishment Order",
    "Duty Letter (For Absent)",
    "Sick Memo",
    "Exam NOC"
])

# Date Input
letter_date = st.date_input("Select Letter Date", date.today())

# Additional Fields
from_date = st.date_input("From Date") if "Duty" in letter_type else None
to_date = st.date_input("To Date") if "Duty" in letter_type else None
duty_date = st.date_input("Join Duty Date") if "Duty" in letter_type else None
memo_text = st.text_area("Memo Text") if "SF-11" in letter_type else ""
exam_name = st.text_input("Exam Name") if "NOC" in letter_type else ""
noc_count = st.selectbox("NOC Attempt No", [1, 2, 3, 4]) if "NOC" in letter_type else None

# Context
context = {
    "LetterDate": letter_date.strftime("%d-%m-%Y"),
    "EmployeeName": selected_emp,
    "Designation": selected_row[17] if len(selected_row) > 17 else "",
    "UnitNumber": selected_row[3] if len(selected_row) > 3 else "",
    "FromDate": from_date.strftime("%d-%m-%Y") if from_date else "",
    "ToDate": to_date.strftime("%d-%m-%Y") if to_date else "",
    "DutyDate": duty_date.strftime("%d-%m-%Y") if duty_date else "",
    "MEMO": memo_text,
    "PFNumber": selected_row[1] if len(selected_row) > 1 else "",
    "ExamName": exam_name,
    "NOCCount": noc_count
}

# Template Mapping
template_files = {
    "SF-11 Punishment Order": "assets/SF-11 Punishment order temp.docx",
    "Duty Letter (For Absent)": "assets/Absent Duty letter temp.docx",
    "Sick Memo": "assets/SICK MEMO temp..docx",
    "Exam NOC": "assets/Exam NOC Letter temp.docx"
}

# Submit Button
if st.button("Generate Letter"):
    docx_path = generate_docx(template_files[letter_type], context)
    st.success("Word letter generated successfully.")
    download_button(docx_path, "⬇️ Download Word Letter")

    pdf_path = convert_to_pdf(docx_path)
    if pdf_path and os.path.exists(pdf_path):
        st.success("PDF letter generated successfully.")
        download_button(pdf_path, "⬇️ Download PDF Letter")
    else:
        st.warning("PDF conversion not supported on this platform. Please use desktop to convert.")
