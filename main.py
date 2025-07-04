import streamlit as st
import pandas as pd
from datetime import date, timedelta
from docx import Document
import os
import base64

st.set_page_config(page_title="Duty Letter Generator", layout="centered")

# Template Mapping
template_path = "assets/Absent Duty letter temp.docx"

# Letter Type Dropdown
st.markdown("## 📌 Duty Letter (For Absent) Generator")

# Load Excel File
employee_master = pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name=None)
sheet_names = list(employee_master.keys())
selected_sheet = st.selectbox("📋 Select Sheet", sheet_names)
df_emp = employee_master[selected_sheet]

# Create Display Column
df_emp["Display"] = df_emp.apply(lambda row: f"{row[1]} - {row[2]} - {row[4]} - {row[5]}", axis=1)
emp_list = df_emp["Display"].dropna().tolist()
selected_display = st.selectbox("👤 Select Employee", emp_list)
emp_row = df_emp[df_emp["Display"] == selected_display].iloc[0]

# Form Inputs
duty_mode = st.selectbox("🛠 Select Mode", ["SF-11 & Duty Letter For Absent", "Duty Letter For Absent"])
from_date = st.date_input("📅 From Date")
to_date = st.date_input("📅 To Date", value=date.today())
join_date = st.date_input("📆 Join Date", value=to_date + timedelta(days=1))
letter_date = st.date_input("📄 Letter Date", value=date.today())

# Extract Employee Details
pf_no = emp_row[1]
hrms_id = emp_row[2]
unit_no = emp_row[4]
working_station = emp_row[8]
emp_eng_name = emp_row[5]
emp_hin_name = emp_row[13]
designation = emp_row[18]
short_name = emp_row[14] if len(emp_row) > 14 else ""

# Letter Number logic (for later use)
letter_no = f"{short_name}/{str(unit_no)[:2]} - {working_station}"

# Auto Memo content
days = (to_date - from_date).days + 1
auto_memo = (
    f"आप बिना किसी पूर्व सूचना के दिनांक {from_date.strftime('%d-%m-%Y')} से "
    f"{to_date.strftime('%d-%m-%Y')} तक कुल {days} दिवस कार्य से अनुपस्थित थे, जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति "
    f"घोर लापरवाही को प्रदर्शित करता है। अतः आप कामों व भूलो के फेहरिस्त धारा 1, 2 एवं 3 के उल्लंघन के दोषी पाए जाते है।"
)

# Placeholder Mapping
context = {
    "LetterDate": letter_date.strftime("%d-%m-%Y"),
    "EmployeeName": emp_hin_name,
    "Designation": designation,
    "FromDate": from_date.strftime("%d-%m-%Y"),
    "ToDate": to_date.strftime("%d-%m-%Y"),
    "JoinDate": join_date.strftime("%d-%m-%Y"),
    "PFNumber": pf_no,
    "Memo": auto_memo,
    "LetterNo": letter_no
}

# Replace Placeholder Function
def generate_filled_doc(template_path, context, emp_name):
    doc = Document(template_path)
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

    filename = f"Duty_Letter_{emp_name.replace(' ', '_')}.docx"
    filepath = os.path.join("generated", filename)
    os.makedirs("generated", exist_ok=True)
    doc.save(filepath)
    return filepath

# Download Button
def download_link(file_path):
    with open(file_path, "rb") as f:
        data = f.read()
    b64 = base64.b64encode(data).decode()
    file_name = os.path.basename(file_path)
    href = f'<a href="data:file/docx;base64,{b64}" download="{file_name}">📥 Download {file_name}</a>'
    st.markdown(href, unsafe_allow_html=True)

# Generate Letter
if st.button("📄 Generate Duty Letter"):
    filled_doc_path = generate_filled_doc(template_path, context, emp_eng_name)
    st.success("✅ Word File Generated Successfully!")
    download_link(filled_doc_path)