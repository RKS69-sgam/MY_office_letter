import streamlit as st
import pandas as pd
import os
import base64
from docx import Document
from datetime import datetime, date, timedelta

# === Output Folder ===
output_folder = "generated_letters"
os.makedirs(output_folder, exist_ok=True)

# === Template Files ===
template_files = {
    "Duty Letter (For Absent)": "assets/Absent Duty letter temp.docx",
    "SF-11 For Other Reason": "assets/SF-11 temp.docx",
    "Sick Memo": "assets/SICK MEMO temp.docx",
    "General Letter": "assets/General Letter temp.docx",
    "Exam NOC": "assets/Exam NOC Letter temp.docx",
    "SF-11 Punishment Order": "assets/SF-11 Punishment order temp.docx"
}

# === Register Paths ===
sf11_register_path = "assets/SF-11 Register.xlsx"
exam_noc_register_path = "assets/ExamNOC_Report.xlsx"

# === Placeholder Replacement ===
def replace_placeholders(doc, context):
    for p in doc.paragraphs:
        inline_replacement(p.runs, context)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                inline_replacement(cell.paragraphs[0].runs, context)

def inline_replacement(runs, context):
    for run in runs:
        for key, val in context.items():
            if f"[{key}]" in run.text:
                run.text = run.text.replace(f"[{key}]", str(val))

# === Generate Word File ===
def generate_word(template_path, context, filename):
    doc = Document(template_path)
    replace_placeholders(doc, context)
    save_path = os.path.join(output_folder, filename)
    doc.save(save_path)
    return save_path

# === Download Link ===
def download_word(path):
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    file_name = os.path.basename(path)
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{file_name}">📥 Download Word File</a>'
    st.markdown(href, unsafe_allow_html=True)

# === Streamlit UI ===
st.title("📄 Railway Letter Generator")

letter_type = st.selectbox("📌 Select Letter Type", [
    "Duty Letter (For Absent)",
    "SF-11 For Other Reason",
    "Sick Memo",
    "General Letter",
    "Exam NOC",
    "SF-11 Punishment Order"
])

# === Load Employee Master Data ===
employee_master = pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name=None)
sheet_names = list(employee_master.keys())
selected_sheet = st.selectbox("📋 Select Sheet", sheet_names)
df_emp = employee_master[selected_sheet]
df_emp["Display"] = df_emp.apply(lambda row: f"{row[1]} - {row[2]} - {row[4]} - {row[5]}", axis=1)
selected_display = st.selectbox("👤 Select Employee", df_emp["Display"].dropna().tolist())
selected_row = df_emp[df_emp["Display"] == selected_display].iloc[0]

# === Extract Info ===
pf_number = selected_row[1]
hrms_id = selected_row[2]
unit = str(selected_row[4])
unit_code = unit[:2]
working_station = selected_row[8]
english_name = selected_row[5]
hindi_name = selected_row[13]
designation = selected_row[18]
short_name = selected_row[14]
letter_no = f"{short_name}/{unit_code}/{working_station}"
letter_date = st.date_input("📄 Letter Date", value=date.today())

# === Letter Type Logic ===

if letter_type == "Duty Letter (For Absent)":
    duty_mode = st.selectbox("🛠 Duty Mode", ["SF-11 & Duty Letter For Absent", "Duty Letter For Absent"])
    from_date = st.date_input("📅 From Date")
    to_date = st.date_input("📅 To Date", value=date.today())
    join_date = st.date_input("📆 Join Date", value=to_date + timedelta(days=1))
    days_absent = (to_date - from_date).days + 1
    memo = f"आप बिना किसी पूर्व सूचना के दिनांक {from_date.strftime('%d-%m-%Y')} से {to_date.strftime('%d-%m-%Y')} तक कुल {days_absent} दिवस कार्य से अनुपस्थित थे, जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है। अतः आप कामों व भूलो के फेहरिस्त धारा 1, 2 एवं 3 के उल्लंघन के दोषी पाए जाते है।"

    context = {
        "LetterDate": letter_date.strftime("%d-%m-%Y"),
        "EmployeeName": hindi_name,
        "Designation": designation,
        "FromDate": from_date.strftime("%d-%m-%Y"),
        "ToDate": to_date.strftime("%d-%m-%Y"),
        "JoinDate": join_date.strftime("%d-%m-%Y"),
        "DutyDate": join_date.strftime("%d-%m-%Y"),
        "PFNumber": pf_number,
        "LetterNo": letter_no,
        "Memo": memo,
        "UnitNumber": unit,
        "ShortName": short_name
    }

elif letter_type == "SF-11 For Other Reason":
    memo_input = st.text_area("📌 Enter Memorandum")
    final_memo = memo_input + " जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है। अतः आप कामों व भूलों के फेहरिस्त धारा 1, 2 एवं 3 के उल्लंघन के दोषी पाए जाते हैं।"

    context = {
        "LetterDate": letter_date.strftime("%d-%m-%Y"),
        "EmployeeName": hindi_name,
        "Designation": designation,
        "PFNumber": pf_number,
        "UnitNumber": unit,
        "ShortName": short_name,
        "Memo": final_memo,
        "LetterNo": letter_no
    }

elif letter_type == "Sick Memo":
    from_date = st.date_input("From Date")
    to_date = st.date_input("To Date")
    join_date = st.date_input("Join Date", value=to_date + timedelta(days=1))
    memo = st.text_area("Remarks")

    context = {
        "LetterDate": letter_date.strftime("%d-%m-%Y"),
        "EmployeeName": hindi_name,
        "Designation": designation,
        "FromDate": from_date.strftime("%d-%m-%Y"),
        "ToDate": to_date.strftime("%d-%m-%Y"),
        "JoinDate": join_date.strftime("%d-%m-%Y"),
        "PFNumber": pf_number,
        "LetterNo": letter_no,
        "Memo": memo,
        "UnitNumber": unit,
        "ShortName": short_name,
        "DutyDate": join_date.strftime("%d-%m-%Y")
    }

elif letter_type == "General Letter":
    subject = st.text_input("Subject")
    ref = st.text_input("Reference (if any)")
    memo = st.text_area("Letter Body / Memo")
    copy_to = st.text_area("Copy To (use commas if multiple)")

    context = {
        "LetterDate": letter_date.strftime("%d-%m-%Y"),
        "EmployeeName": hindi_name,
        "Designation": designation,
        "PFNumber": pf_number,
        "LetterNo": letter_no,
        "Memo": memo,
        "UnitNumber": unit,
        "ShortName": short_name,
        "Subject": subject,
        "Reference": ref,
        "CopyTo": copy_to
    }

elif letter_type == "Exam NOC":
    exam_name = st.text_input("Exam Name")
    noc_year = st.selectbox("Select Year", [2023, 2024, 2025])
    application_no = st.selectbox("Application Count (1–4)", [1, 2, 3, 4])
    memo = f"Permission is granted to appear in {exam_name} for the year {noc_year} (Application #{application_no})."

    context = {
        "LetterDate": letter_date.strftime("%d-%m-%Y"),
        "EmployeeName": hindi_name,
        "Designation": designation,
        "PFNumber": pf_number,
        "LetterNo": letter_no,
        "Memo": memo,
        "UnitNumber": unit,
        "ShortName": short_name
    }

elif letter_type == "SF-11 Punishment Order":
    punishment = st.text_area("Enter Punishment Order / Memo")

    context = {
        "LetterDate": letter_date.strftime("%d-%m-%Y"),
        "EmployeeName": hindi_name,
        "Designation": designation,
        "PFNumber": pf_number,
        "LetterNo": letter_no,
        "Memo": punishment,
        "UnitNumber": unit,
        "ShortName": short_name
    }

# === Generate ===
if st.button("📄 Generate Letter"):
    template_path = template_files[letter_type]
    filename = f"{letter_type} - {hindi_name}.docx"
    filepath = generate_word(template_path, context, filename)
    st.success("✅ Letter generated successfully!")
    download_word(filepath)