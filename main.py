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

# === Template Mapping ===
template_files = {
    "Duty Letter (For Absent)": "assets/Absent Duty letter temp.docx",
    "SF-11 For Other Reason": "assets/SF-11 temp.docx",
    "Sick Memo": "assets/Sick Memo temp.docx",
    "General Letter": "assets/General Letter temp.docx",
    "Exam NOC": "assets/Exam NOC temp.docx",
    "SF-11 Punishment Order": "assets/SF-11 Punishment temp.docx"
}

# === Placeholder Replacement Function ===
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

# === Word Generation Function ===
def generate_word(template_path, context, filename):
    doc = Document(template_path)
    replace_placeholder_runs(doc, context)
    save_path = os.path.join(output_folder, filename)
    doc.save(save_path)
    return save_path

# === Download Link Function ===
def download_word(path):
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    file_name = os.path.basename(path)
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{file_name}">📥 Download Word File</a>'
    st.markdown(href, unsafe_allow_html=True)

# === UI ===
st.title("📄 Railway Letter Generator")

letter_type = st.selectbox("📌 Select Letter Type:", list(template_files.keys()))

# === Load Employee Master Data ===
employee_master = pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name=None)
sheet_names = list(employee_master.keys())
selected_sheet = st.selectbox("📋 Select Sheet", sheet_names)
df_emp = employee_master[selected_sheet]
df_emp["Display"] = df_emp.apply(lambda row: f"{row[1]} - {row[2]} - {row[4]} - {row[5]}", axis=1)
selected_display = st.selectbox("👤 Select Employee", df_emp["Display"].dropna().tolist())
selected_row = df_emp[df_emp["Display"] == selected_display].iloc[0]

# === Employee Info ===
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

context = {
    "EmployeeName": hindi_name,
    "Designation": designation,
    "PFNumber": pf_number,
    "UnitNumber": unit,
    "ShortName": short_name,
    "LetterNo": letter_no
}

# === Common Date Input ===
letter_date = st.date_input("📄 Letter Date", value=date.today())
context["LetterDate"] = letter_date.strftime("%d-%m-%Y")

# === Duty Letter ===
if letter_type == "Duty Letter (For Absent)":
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

# === SF-11 For Other Reason ===
elif letter_type == "SF-11 For Other Reason":
    user_memo = st.text_area("📌 Enter Memorandum")
    full_memo = user_memo.strip() + " जो कि रेलवे सेवा के प्रति घोर लापरवाही को प्रदर्शित करता है।"
    context["Memo"] = full_memo

# === Other letters can also be handled similarly using context inputs ===

# === Generate Button ===
if st.button("📄 Generate Letter"):
    template_path = template_files[letter_type]
    filename = f"{letter_type} - {hindi_name}.docx"
    file_path = generate_word(template_path, context, filename)
    st.success("✅ Letter generated successfully!")
    download_word(file_path)