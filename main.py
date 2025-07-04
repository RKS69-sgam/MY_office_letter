import streamlit as st
import pandas as pd
import os
import base64
from docx import Document
from datetime import datetime, date, timedelta

# === Output Folder ===
output_folder = "generated_letters"
os.makedirs(output_folder, exist_ok=True)

# === Define Template Files ===
template_files = {
    "Duty Letter (For Absent)": "assets/Absent Duty letter temp.docx",
    "SF-11 For Other Reason": "assets/SF-11 temp.docx",
    "Sick Memo": "assets/SICK MEMO temp.docx",
    "General Letter": "assets/General Letter temp.docx",
    "Exam NOC": "assets/Exam NOC Letter temp.docx",
    "SF-11 Punishment Order": "assets/SF-11 Punishment order temp.docx"
}

# === Function to Generate Word Document ===
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

# === Function to Download Word File ===
def download_word(path):
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    file_name = os.path.basename(path)
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{file_name}">\ud83d\udcc5 Download Word File</a>'
    st.markdown(href, unsafe_allow_html=True)

# === Streamlit UI ===
st.title("\ud83d\udcc4 Railway Letter Generator")

letter_type = st.selectbox("\ud83d\udd39 Select Letter Type:", list(template_files.keys()))

# === Load Employee Data ===
employee_master = pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name=None)
sheet_names = list(employee_master.keys())
selected_sheet = st.selectbox("\ud83d\udccb Select Sheet", sheet_names)
df_emp = employee_master[selected_sheet]
df_emp["Display"] = df_emp.apply(lambda row: f"{row[1]} - {row[2]} - {row[4]} - {row[5]}", axis=1)
selected_display = st.selectbox("\ud83d\udc64 Select Employee", df_emp["Display"].dropna().tolist())
selected_row = df_emp[df_emp["Display"] == selected_display].iloc[0]

# === Extract Common Fields ===
pf_number = selected_row[1]
hindi_name = selected_row[13]
designation = selected_row[18]
short_name = selected_row[14]
unit = str(selected_row[4])
working_station = selected_row[8]

# === Letter Specific Input ===
letter_date = st.date_input("\ud83d\udcc4 Letter Date", value=date.today())

# === Context Template ===
context = {
    "LetterDate": letter_date.strftime("%d-%m-%Y"),
    "EmployeeName": hindi_name,
    "Designation": designation,
    "PFNumber": pf_number,
    "UnitNumber": unit,
    "ShortName": short_name,
    "LetterNo": f"{short_name}/{unit[:2]}/{working_station}",
    "Memo": "",
    "FromDate": "",
    "ToDate": "",
    "JoinDate": "",
    "DutyDate": ""
}

# === Duty Letter Handling ===
if letter_type == "Duty Letter (For Absent)":
    duty_mode = st.selectbox("\ud83d\udee0 Duty Mode", ["SF-11 & Duty Letter For Absent", "Duty Letter For Absent"])
    from_date = st.date_input("\ud83d\uddd3\ufe0f From Date")
    to_date = st.date_input("\ud83d\uddd3\ufe0f To Date", value=date.today())
    join_date = st.date_input("\ud83d\udcc6 Join Date", value=to_date + timedelta(days=1))

    context.update({
        "FromDate": from_date.strftime("%d-%m-%Y"),
        "ToDate": to_date.strftime("%d-%m-%Y"),
        "JoinDate": join_date.strftime("%d-%m-%Y"),
        "DutyDate": join_date.strftime("%d-%m-%Y"),
        "Memo": f"आप बिना किसी पूर्व सूचना के दिनांक {from_date.strftime('%d-%m-%Y')} से {to_date.strftime('%d-%m-%Y')} तक कुल {(to_date - from_date).days + 1} दिवस कार्य से अनुपस्थित थे, जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है। अतः आप कामों व भूलो के फेहरिस्त धारा 1, 2 एवं 3 के उल्लंघन के दोषी पाए जाते है।"
    })

    if st.button("\ud83d\udcc4 Generate Duty Letter"):
        if duty_mode == "SF-11 & Duty Letter For Absent":
            sf11_path = generate_word(template_files["SF-11 For Other Reason"], context, f"SF-11 - {hindi_name}.docx")
            st.success("\u2705 SF-11 Letter generated!")
            download_word(sf11_path)
        duty_path = generate_word(template_files["Duty Letter (For Absent)"], context, f"Duty Letter - {hindi_name}.docx")
        st.success("\u2705 Duty Letter generated!")
        download_word(duty_path)

elif letter_type == "SF-11 For Other Reason":
    memo = st.text_area("\ud83d\udcc4 Memorandum")
    context["Memo"] = memo + " जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है।"
    if st.button("\ud83d\udcc4 Generate SF-11 Letter"):
        sf11_path = generate_word(template_files["SF-11 For Other Reason"], context, f"SF-11 - {hindi_name}.docx")
        st.success("\u2705 SF-11 Letter generated!")
        download_word(sf11_path)

elif letter_type in template_files:
    if st.button("\ud83d\udcc4 Generate Letter"):
        file_path = generate_word(template_files[letter_type], context, f"{letter_type} - {hindi_name}.docx")
        st.success(f"\u2705 {letter_type} generated!")
        download_word(file_path)
