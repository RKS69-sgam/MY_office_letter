import streamlit as st
import pandas as pd
import os
import base64
from docx import Document
from datetime import date, timedelta

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

# === Letter Type Dropdown ===
letter_types = list(template_files.keys())
selected_letter_type = st.selectbox("📌 Select Letter Type:", letter_types)

# === Load Employee Master Data ===
employee_master = pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name=None)
sheet_names = list(employee_master.keys())
selected_sheet = st.selectbox("📋 Select Sheet", sheet_names)
df_emp = employee_master[selected_sheet]
df_emp["Display"] = df_emp.apply(lambda row: f"{row[1]} - {row[2]} - {row[4]} - {row[5]}", axis=1)
emp_display_list = df_emp["Display"].dropna().tolist()
selected_emp_display = st.selectbox("👤 Select Employee:", emp_display_list)
selected_row = df_emp[df_emp["Display"] == selected_emp_display].iloc[0]

# === Extract Employee Info ===
pf_number = selected_row[1]
hrms_id = selected_row[2]
unit = str(selected_row[4])
working_station = selected_row[8]
english_name = selected_row[5]
hindi_name = selected_row[13]
designation = selected_row[18]
short_name = selected_row[14] if len(selected_row) > 14 else ""

# === Common Date Fields ===
letter_date = st.date_input("📄 Letter Date", value=date.today())
from_date = st.date_input("📅 From Date")
to_date = st.date_input("📅 To Date", value=date.today())
join_date = st.date_input("📆 Join Date", value=to_date + timedelta(days=1))

# === Letter No. and Memo ===
unit_code = unit[:2] if len(unit) >= 2 else unit
letter_no = f"{short_name}/{unit_code}/{working_station}"
days_absent = (to_date - from_date).days + 1
memo = f"आप बिना किसी पूर्व सूचना के दिनांक {from_date.strftime('%d-%m-%Y')} से {to_date.strftime('%d-%m-%Y')} तक कुल {days_absent} दिवस कार्य से अनुपस्थित थे, जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है। अतः आप कामों व भूलो के फेहरिस्त धारा 1, 2 एवं 3 के उल्लंघन के दोषी पाए जाते है।"

# === Placeholder Mapping ===
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
    "UnitNumber": unit
}

# === Generate Word File ===
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
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{file_name}">📥 Download Word File</a>'
    st.markdown(href, unsafe_allow_html=True)

# === Duty Letter Special Case ===
if selected_letter_type == "Duty Letter (For Absent)":
    duty_mode = st.selectbox("🛠 Duty Letter Mode", [
        "SF-11 & Duty Letter For Absent",
        "Duty Letter For Absent"
    ])

    if st.button("📄 Generate Letter"):
        if duty_mode == "SF-11 & Duty Letter For Absent":
            # SF-11
            sf11_template = template_files["SF-11 For Other Reason"]
            sf11_filename = f"SF-11 - {hindi_name}.docx"
            sf11_path = generate_word(sf11_template, context, sf11_filename)
            st.success("✅ SF-11 Letter generated successfully!")
            download_word(sf11_path)

            # Duty Letter
            duty_template = template_files["Duty Letter (For Absent)"]
            duty_filename = f"Duty Letter - {hindi_name}.docx"
            duty_path = generate_word(duty_template, context, duty_filename)
            st.success("✅ Duty Letter generated successfully!")
            download_word(duty_path)

        elif duty_mode == "Duty Letter For Absent":
            duty_template = template_files["Duty Letter (For Absent)"]
            duty_filename = f"Duty Letter - {hindi_name}.docx"
            duty_path = generate_word(duty_template, context, duty_filename)
            st.success("✅ Duty Letter generated successfully!")
            download_word(duty_path)

# === Other Letters ===
else:
    if st.button("📄 Generate Letter"):
        template_path = template_files[selected_letter_type]
        filename = f"{selected_letter_type} - {hindi_name}.docx"
        filled_path = generate_word(template_path, context, filename)
        st.success(f"✅ {selected_letter_type} generated successfully!")
        download_word(filled_path)