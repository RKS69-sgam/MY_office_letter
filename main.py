import streamlit as st
import pandas as pd
from datetime import date, timedelta
from docx import Document
from tempfile import NamedTemporaryFile
import base64
import os

# Load Excel
@st.cache_data
def load_employee_data():
    return pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name=None)

# Replace placeholders
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

# Generate .docx
def generate_docx(template_path, context):
    doc = Document(template_path)
    replace_placeholders(doc, context)
    temp_file = NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(temp_file.name)
    return temp_file.name

# Convert to PDF (optional)
def convert_to_pdf(docx_path):
    try:
        from docx2pdf import convert
        pdf_path = docx_path.replace(".docx", ".pdf")
        convert(docx_path, pdf_path)
        return pdf_path
    except:
        return None

# Download button
def download_button(file_path, label):
    with open(file_path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
        href = f'<a href="data:application/octet-stream;base64,{b64}" download="{os.path.basename(file_path)}">{label}</a>'
        st.markdown(href, unsafe_allow_html=True)

# Format Unit Number
def format_unit(unit_val, station_val):
    try:
        if any(char.isalpha() for char in str(unit_val)):
            return f"{unit_val}/{station_val}"
        else:
            return f"{str(unit_val)[:2]}/{station_val}"
    except:
        return unit_val

# Load Data
data = load_employee_data()
sheet_names = list(data.keys())
selected_sheet = st.selectbox("Select Unit Sheet:", sheet_names)
df = data[selected_sheet]

# Column mappings
col_pf = 1
col_hrms = 2
col_unit = 4
col_eng_name = 5
col_hindi_name = 13
col_designation = 18
col_station = 8

# Dropdown string
df["Dropdown"] = df.apply(
    lambda row: f"(PF:{row[col_pf]}, HRMS:{row[col_hrms]}, Unit:{format_unit(row[col_unit], row[col_station])}) {row[col_eng_name]}",
    axis=1
)

dropdown_list = ["-- Select Employee --"] + df["Dropdown"].tolist()
selected_dropdown = st.selectbox("Select Employee (with details):", dropdown_list)

if selected_dropdown == "-- Select Employee --":
    st.warning("Please select an employee to proceed.")
    st.stop()
else:
    selected_row = df[df["Dropdown"] == selected_dropdown].iloc[0]

# Letter Type
letter_type = st.selectbox("Select Letter Type:", [
    "SF-11 Punishment Order",
    "Duty Letter (For Absent)",
    "Sick Memo",
    "Exam NOC"
])

letter_date = st.date_input("Select Letter Date", date.today())

from_date = st.date_input("From Date") if "Duty" in letter_type else None
to_date = st.date_input("To Date") if "Duty" in letter_type else None

# Join Date = To Date + 1 (editable)
duty_date = st.date_input("Join Duty Date", to_date + timedelta(days=1) if to_date else date.today()) if "Duty" in letter_type else None

memo_text = st.text_area("Memo Text") if "SF-11" in letter_type else ""
exam_name = st.text_input("Exam Name") if "NOC" in letter_type else ""
noc_count = st.selectbox("NOC Attempt No", [1, 2, 3, 4]) if "NOC" in letter_type else None

# Context dictionary
context = {
    "LetterDate": letter_date.strftime("%d-%m-%Y"),
    "EmployeeName": selected_row[col_hindi_name],
    "Designation": selected_row[col_designation] if len(selected_row) > col_designation else "",
    "UnitNumber": format_unit(selected_row[col_unit], selected_row[col_station]),
    "FromDate": from_date.strftime("%d-%m-%Y") if from_date else "",
    "ToDate": to_date.strftime("%d-%m-%Y") if to_date else "",
    "DutyDate": duty_date.strftime("%d-%m-%Y") if duty_date else "",
    "MEMO": memo_text,
    "PFNumber": selected_row[col_pf],
    "ExamName": exam_name,
    "NOCCount": noc_count
}

# Template Paths
template_files = {
    "SF-11 Punishment Order": "assets/SF-11 Punishment order temp.docx",
    "Duty Letter (For Absent)": "assets/Absent Duty letter temp.docx",
    "Sick Memo": "assets/SICK MEMO temp..docx",
    "Exam NOC": "assets/Exam NOC Letter temp.docx"
}

# Generate
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