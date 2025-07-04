import streamlit as st
import pandas as pd
import os
import base64
from docx import Document
from datetime import datetime, timedelta

# === Output Folder ===
output_folder = "generated_letters"
os.makedirs(output_folder, exist_ok=True)

# === Template Files Path ===
template_files = {
    "Duty Letter (For Absent)": "assets/Absent Duty letter temp.docx",
    "SF-11 For Other Reason": "assets/SF-11 temp.docx",
    "Sick Memo": "assets/SICK MEMO temp..docx",
    "Exam NOC": "assets/Exam NOC Letter temp.docx",
    "General Letter": "assets/General Letter temp.docx",
    "SF-11 Punishment Order": "assets/SF-11 Punishment order temp.docx"
}

# === UI ===
st.title("📄 Generate Duty / SF-11 / Memo Letter")

letter_type = st.selectbox("📌 Select Letter Type:", list(template_files.keys()))

# === Load Excel Master Data ===
employee_master = pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name=None)
sheet_names = list(employee_master.keys())
selected_sheet = st.selectbox("📋 Select Sheet", sheet_names)
df_emp = employee_master[selected_sheet]
df_emp["Display"] = df_emp.apply(lambda row: f"{row[1]} - {row[2]} - {row[4]} - {row[5]}", axis=1)
selected_display = st.selectbox("👤 Select Employee", df_emp["Display"].dropna().tolist())
selected_row = df_emp[df_emp["Display"] == selected_display].iloc[0]

# === Employee Fields ===
pf_number = selected_row[1]
emp_name = selected_row[13]
designation = selected_row[18]
short_name = selected_row[14]
unit = str(selected_row[4])
station = selected_row[8]

# === Common Date Fields ===
letter_date = st.date_input("📄 Letter Date", datetime.today())
from_date = st.date_input("📅 From Date")
to_date = st.date_input("📅 To Date")
join_date = st.date_input("📆 Join Date", value=to_date + timedelta(days=1))

# === Memo Field for SF-11 (Other Reason only) ===
memo = ""
if letter_type == "SF-11 For Other Reason":
    memo = st.text_area("📝 Enter Memo for SF-11", placeholder="यहां हिंदी में मेमो टाइप करें...")
else:
    days_absent = (to_date - from_date).days + 1
    memo = f"आप बिना किसी पूर्व सूचना के दिनांक {from_date.strftime('%d-%m-%Y')} से {to_date.strftime('%d-%m-%Y')} तक कुल {days_absent} दिवस कार्य से अनुपस्थित थे..."

# === Placeholder Context ===
context = {
    "LetterDate": letter_date.strftime("%d-%m-%Y"),
    "FromDate": from_date.strftime("%d-%m-%Y"),
    "ToDate": to_date.strftime("%d-%m-%Y"),
    "JoinDate": join_date.strftime("%d-%m-%Y"),
    "DutyDate": join_date.strftime("%d-%m-%Y"),
    "EmployeeName": emp_name,
    "Designation": designation,
    "PFNumber": pf_number,
    "Unit": unit,
    "ShortName": short_name,
    "Memo": memo,
    "LetterNo": f"{short_name}/{unit[:2]}/{station}",
    "UnitNumber": unit
}

# === Word File Generation ===
def generate_word(template_path, context, filename):
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
    save_path = os.path.join(output_folder, filename)
    doc.save(save_path)
    return save_path

# === Download Link ===
def download_word(path):
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    file_name = os.path.basename(path)
    st.markdown(f'<a href="data:application/octet-stream;base64,{b64}" download="{file_name}">📥 Download Word File</a>', unsafe_allow_html=True)

# === Generate Button ===
if st.button("📄 Generate Letter"):
    template_path = template_files[letter_type]
    filename = f"{letter_type} - {emp_name}.docx"
    save_path = generate_word(template_path, context, filename)
    st.success(f"✅ Letter generated: {filename}")
    download_word(save_path)