import streamlit as st
import pandas as pd
from datetime import date, timedelta
from docx import Document
import base64
import os

# === Load Letter Type first ===
letter_type = st.selectbox("Select Letter Type:", [
    "SF-11 Punishment Order",
    "Duty Letter (For Absent)",
    "Sick Memo",
    "Exam NOC"
])

# === Load Excel sources ===
@st.cache_data
def load_excel_data():
    emp_data = pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name=None)
    sf11_data = pd.read_excel("assets/SF-11 Register.xlsx", sheet_name=None)
    return emp_data, sf11_data

emp_data, sf11_data = load_excel_data()

# === Select Sheet (common for both files) ===
sheet_names = list(emp_data.keys())
selected_sheet = st.selectbox("Select Unit Sheet:", sheet_names)

# === Conditional source selection ===
if letter_type == "SF-11 Punishment Order":
    df = sf11_data[selected_sheet]
    df["Dropdown"] = df.apply(lambda row: f"{row[2]} ({row[6]})", axis=1)  # EmployeeName (LetterNo)
else:
    df = emp_data[selected_sheet]
    def correct_unit_station_format(unit_val, station_val):
        try:
            unit_str = str(unit_val).strip()
            if "/" in unit_str:
                unit_str = unit_str.split("/")[0].strip()
            if unit_str.isdigit():
                unit_str = unit_str[:2]
            station_str = str(station_val).strip()
            return f"{unit_str}/{station_str}"
        except:
            return ""
    df["DisplayUnit"] = df.apply(lambda row: correct_unit_station_format(row[4], row["WORKING STATION"]), axis=1)
    df["Dropdown"] = df.apply(lambda row: f"{row[1]} - {row[2]} - {row['DisplayUnit']} - {row[5]}", axis=1)

# === Select Employee ===
selected_dropdown = st.selectbox("Select Employee:", df["Dropdown"].dropna().tolist())
selected_row = df[df["Dropdown"] == selected_dropdown].iloc[0]

# === Letter Specific Fields ===
letter_date = st.date_input("Select Letter Date", date.today())
from_date = st.date_input("From Date") if "Duty" in letter_type else None
to_date = st.date_input("To Date") if "Duty" in letter_type else None
duty_date_default = (to_date + timedelta(days=1)) if to_date else date.today()
duty_date = st.date_input("Join Duty Date", duty_date_default) if "Duty" in letter_type else None
memo_text = st.text_area("Memo Text") if "SF-11" in letter_type else ""
exam_name = st.text_input("Exam Name") if "NOC" in letter_type else ""
noc_count = st.selectbox("NOC Attempt No", [1, 2, 3, 4]) if "NOC" in letter_type else None

# === Context Preparation ===
context = {
    "LetterDate": letter_date.strftime("%d-%m-%Y"),
    "EmployeeName": selected_row[13] if letter_type != "SF-11 Punishment Order" else selected_row[2],
    "Designation": selected_row[18] if letter_type != "SF-11 Punishment Order" else "",
    "UnitNumber": selected_row["DisplayUnit"] if letter_type != "SF-11 Punishment Order" else "",
    "FromDate": from_date.strftime("%d-%m-%Y") if from_date else "",
    "ToDate": to_date.strftime("%d-%m-%Y") if to_date else "",
    "DutyDate": duty_date.strftime("%d-%m-%Y") if duty_date else "",
    "MEMO": memo_text,
    "PFNumber": selected_row[1],
    "ExamName": exam_name,
    "NOCCount": noc_count
}

# === Template Files ===
template_files = {
    "SF-11 Punishment Order": "assets/SF-11 Punishment order temp.docx",
    "Duty Letter (For Absent)": "assets/Absent Duty letter temp.docx",
    "Sick Memo": "assets/SICK MEMO temp..docx",
    "Exam NOC": "assets/Exam NOC Letter temp.docx"
}

# === Replace placeholders ===
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

# === Generate docx ===
def generate_docx(template_path, context, filename):
    doc = Document(template_path)
    replace_placeholders(doc, context)
    docx_path = os.path.join("/tmp", filename + ".docx")
    doc.save(docx_path)
    return docx_path

# === Download Link ===
def download_button(file_path, label):
    with open(file_path, "rb") as f:
        data = f.read()
        b64 = base64.b64encode(data).decode()
        href = f'<a href="data:application/octet-stream;base64,{b64}" download="{os.path.basename(file_path)}">{label}</a>'
        st.markdown(href, unsafe_allow_html=True)

# === Final Button ===
if st.button("Generate Letter"):
    base_filename = f"{letter_type.split()[0]}_{context['EmployeeName']}_{letter_date.strftime('%d-%m-%Y')}"
    docx_path = generate_docx(template_files[letter_type], context, base_filename)
    st.success("Word letter generated successfully.")
    download_button(docx_path, f"⬇️ Download {os.path.basename(docx_path)}")