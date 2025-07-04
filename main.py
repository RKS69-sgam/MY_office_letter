import streamlit as st
import pandas as pd
import os
import base64
from docx import Document
from datetime import datetime, date, timedelta

# === Output Folder ===
output_folder = "generated_letters"
os.makedirs(output_folder, exist_ok=True)

# === Template Mapping ===
template_files = {
    "Duty Letter (For Absent)": "assets/Absent Duty letter temp.docx",
    "SF-11 For Other Reason": "assets/SF-11 temp.docx",
    "Sick Memo": "assets/SICK MEMO temp.docx",
    "General Letter": "assets/General Letter temp.docx",
    "Exam NOC": "assets/Exam NOC Letter temp.docx",
    "SF-11 Punishment Order": "assets/SF-11 Punishment order temp.docx"
}

# === Streamlit UI ===
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

# === Extract Employee Info ===
pf_number = selected_row[1]
hrms_id = selected_row[2]
unit = str(selected_row[4])
unit_code = unit[:2] if len(unit) >= 2 else unit
working_station = selected_row[8]
english_name = selected_row[5]
hindi_name = selected_row[13]
designation = selected_row[18]
short_name = selected_row[14]
letter_no = f"{short_name}/{unit_code}/{working_station}"

# === Inputs common or specific ===
letter_date = st.date_input("📅 Letter Date", value=date.today())

if letter_type == "Duty Letter (For Absent)":
    duty_mode = st.selectbox("🛠 Duty Mode", ["SF-11 & Duty Letter For Absent", "Duty Letter For Absent"])
    from_date = st.date_input("From Date")
    to_date = st.date_input("To Date", value=date.today())
    join_date = st.date_input("Join Date", value=to_date + timedelta(days=1))
    days_absent = (to_date - from_date).days + 1
    memo = f"आप बिना किसी पूर्व सूचना के दिनांक {from_date.strftime('%d-%m-%Y')} से {to_date.strftime('%d-%m-%Y')} तक कुल {days_absent} दिवस कार्य से अनुपस्थित थे, जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है। अतः आप कामों व भूलो के फेहरिस्त धारा 1, 2 एवं 3 के उल्लंघन के दोषी पाए जाते है।"

    context = {
        "LetterDate": letter_date.strftime("%d-%m-%Y"),
        "EmployeeName": hindi_name,
        "Designation": designation,
        "FromDate": from_date.strftime("%d-%m-%Y"),
        "ToDate": to_date.strftime("%d-%m-%Y"),
        "JoinDate": join_date.strftime("%d-%m-%Y"),
        "PFNumber": pf_number,
        "UnitNumber": unit,
        "ShortName": short_name,
        "LetterNo": letter_no,
        "Memo": memo,
        "DutyDate": join_date.strftime("%d-%m-%Y")
    }

elif letter_type == "SF-11 For Other Reason":
    memo_input = st.text_area("📌 Enter Memo")
    final_memo = memo_input.strip() + " जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है।"
    context = {
        "LetterDate": letter_date.strftime("%d-%m-%Y"),
        "EmployeeName": hindi_name,
        "Designation": designation,
        "PFNumber": pf_number,
        "UnitNumber": unit,
        "ShortName": short_name,
        "LetterNo": letter_no,
        "Memo": final_memo
    }

else:
    memo = st.text_area("📌 Enter Memo")
    context = {
        "LetterDate": letter_date.strftime("%d-%m-%Y"),
        "EmployeeName": hindi_name,
        "Designation": designation,
        "PFNumber": pf_number,
        "UnitNumber": unit,
        "ShortName": short_name,
        "LetterNo": letter_no,
        "Memo": memo
    }

# === Function: Run-safe placeholder replace ===
def replace_placeholder_runs(doc, context):
    from docx.oxml.text.paragraph import CT_P
    from docx.text.paragraph import Paragraph
    from docx.oxml.text.run import CT_R
    from docx.text.run import Run

    def replace_in_runs(runs, key, val):
        for run in runs:
            if key in run.text:
                run.text = run.text.replace(key, val)

    for para in doc.paragraphs:
        full_text = "".join(run.text for run in para.runs)
        for key, val in context.items():
            placeholder = f"[{key}]"
            if placeholder in full_text:
                replace_in_runs(para.runs, placeholder, str(val))

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    full_text = "".join(run.text for run in para.runs)
                    for key, val in context.items():
                        placeholder = f"[{key}]"
                        if placeholder in full_text:
                            replace_in_runs(para.runs, placeholder, str(val))

# === Generate and Download ===
def generate_and_download(template_key, filename):
    doc = Document(template_files[template_key])
    replace_placeholder_runs(doc, context)
    save_path = os.path.join(output_folder, filename)
    doc.save(save_path)
    with open(save_path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{filename}">📥 Download {filename}</a>'
    st.markdown(href, unsafe_allow_html=True)

# === Generate Button ===
if st.button("📄 Generate Letter"):
    if letter_type == "Duty Letter (For Absent)":
        if duty_mode == "SF-11 & Duty Letter For Absent":
            generate_and_download("SF-11 For Other Reason", f"SF-11 - {hindi_name}.docx")
        generate_and_download("Duty Letter (For Absent)", f"Duty Letter - {hindi_name}.docx")

    else:
        generate_and_download(letter_type, f"{letter_type} - {hindi_name}.docx")