import streamlit as st
import pandas as pd
import os
import base64
from docx import Document
from datetime import datetime, date, timedelta
from docx.text.paragraph import Paragraph

# === Output Folder ===
output_folder = "generated_letters"
os.makedirs(output_folder, exist_ok=True)

# === Templates Map ===
template_files = {
    "Duty Letter (For Absent)": "assets/Absent Duty letter temp.docx",
    "SF-11 For Other Reason": "assets/SF-11 temp.docx",
    "Sick Memo": "assets/SICK MEMO temp.docx",
    "Exam NOC": "assets/Exam NOC Letter temp.docx",
    "General Letter": "assets/General Letter temp.docx"
}

# === Replace Placeholders (including table runs) ===
def replace_placeholder_runs(doc, context):
    def process_paragraph(paragraph: Paragraph):
        full_text = ''.join(run.text for run in paragraph.runs)
        for key, val in context.items():
            if f"[{key}]" in full_text:
                full_text = full_text.replace(f"[{key}]", str(val))
                for run in paragraph.runs:
                    run.text = ''
                if paragraph.runs:
                    paragraph.runs[0].text = full_text

    for para in doc.paragraphs:
        process_paragraph(para)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    process_paragraph(para)

# === Generate Word ===
def generate_word(template_path, context, filename):
    doc = Document(template_path)
    replace_placeholder_runs(doc, context)
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

# === UI ===
st.title("📄 Railway Letter Generator")

letter_type = st.selectbox("📌 Select Letter Type:", list(template_files.keys()) + ["SF-11 & Duty Letter For Absent"])

# === Employee Data ===
employee_master = pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name=None)
sheet_names = list(employee_master.keys())
selected_sheet = st.selectbox("📋 Select Sheet", sheet_names)
df_emp = employee_master[selected_sheet]
df_emp["Display"] = df_emp.apply(lambda row: f"{row[1]} - {row[2]} - {row[4]} - {row[5]}", axis=1)
selected_display = st.selectbox("👤 Select Employee", df_emp["Display"].dropna().tolist())
selected_row = df_emp[df_emp["Display"] == selected_display].iloc[0]

# === Extract Fields ===
pf_number = selected_row[1]
hrms_id = selected_row[2]
unit = str(selected_row[4])
working_station = selected_row[8]
english_name = selected_row[5]
hindi_name = selected_row[13]
designation = selected_row[18]
short_name = selected_row[14]
unit_code = unit[:2] if len(unit) >= 2 else unit
letter_no = f"{short_name}/{unit_code}/{working_station}"

# === Date Fields Common ===
letter_date = st.date_input("📄 Letter Date", value=date.today())

# === Context Init ===
context = {
    "LetterDate": letter_date.strftime("%d-%m-%Y"),
    "EmployeeName": hindi_name,
    "Designation": designation,
    "PFNumber": pf_number,
    "UnitNumber": unit,
    "ShortName": short_name,
    "LetterNo": letter_no
}

# === Letter Specific Inputs ===
if letter_type == "Duty Letter (For Absent)" or letter_type == "SF-11 & Duty Letter For Absent":
    from_date = st.date_input("📅 From Date")
    to_date = st.date_input("📅 To Date", value=date.today())
    join_date = st.date_input("📆 Join Date", value=to_date + timedelta(days=1))

    context["FromDate"] = from_date.strftime("%d-%m-%Y")
    context["ToDate"] = to_date.strftime("%d-%m-%Y")
    context["JoinDate"] = join_date.strftime("%d-%m-%Y")
    context["DutyDate"] = join_date.strftime("%d-%m-%Y")

    days_absent = (to_date - from_date).days + 1
    memo = f"आप बिना किसी पूर्व सूचना के दिनांक {from_date.strftime('%d-%m-%Y')} से {to_date.strftime('%d-%m-%Y')} तक कुल {days_absent} दिवस कार्य से अनुपस्थित थे, जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है। अतः आप कामों व भूलो के फेहरिस्त धारा 1, 2 एवं 3 के उल्लंघन के दोषी पाए जाते है।"
    context["Memo"] = memo

elif letter_type == "SF-11 For Other Reason":
    memo = st.text_area("📌 Enter Memorandum")
    context["Memo"] = memo + "\nजो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है।"

elif letter_type == "Sick Memo" or letter_type == "General Letter":
    memo = st.text_area("📌 Remarks / Memo")
    context["Memo"] = memo

elif letter_type == "Exam NOC":
    exam_name = st.text_input("📘 Exam Name")
    context["Memo"] = f"Exam NOC requested for: {exam_name}"

# === Generate Button ===
if st.button("📄 Generate Letter"):
    if letter_type == "SF-11 & Duty Letter For Absent":
        sf_template = template_files["SF-11 For Other Reason"]
        duty_template = template_files["Duty Letter (For Absent)"]
        sf_file = generate_word(sf_template, context, f"SF-11 - {hindi_name}.docx")
        duty_file = generate_word(duty_template, context, f"Duty Letter - {hindi_name}.docx")
        st.success("✅ SF-11 and Duty Letter generated!")
        download_word(sf_file)
        download_word(duty_file)

    else:
        template_path = template_files.get(letter_type)
        filename = f"{letter_type} - {hindi_name}.docx"
        path = generate_word(template_path, context, filename)
        st.success("✅ Letter generated!")
        download_word(path)