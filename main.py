# === Streamlit Letter Generator App ===

import streamlit as st
import pandas as pd
import os
import base64
from docx import Document
from datetime import datetime, date, timedelta

# === Directories ===
output_folder = "generated_letters"
os.makedirs(output_folder, exist_ok=True)

# === Templates ===
template_files = {
    "Duty Letter (For Absent)": "assets/Absent Duty letter temp.docx",
    "SF-11 For Other Reason": "assets/SF-11 temp.docx",
    "Sick Memo": "assets/SICK MEMO temp.docx",
    "General Letter": "assets/General Letter temp.docx",
    "Exam NOC": "assets/Exam NOC Letter temp.docx",
    "SF-11 Punishment Order": "assets/SF-11 Punishment order temp.docx"
}

# === Registers ===
sf11_register_path = "assets/SF-11 Register.xlsx"
exam_noc_path = "assets/ExamNOC_Report.xlsx"

# === Helper Function: Placeholder Replace ===
def generate_doc(template_path, context, filename):
    doc = Document(template_path)
    for p in doc.paragraphs:
        for run in p.runs:
            for key, val in context.items():
                if f"[{key}]" in run.text:
                    run.text = run.text.replace(f"[{key}]", str(val))
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        for key, val in context.items():
                            if f"[{key}]" in run.text:
                                run.text = run.text.replace(f"[{key}]", str(val))
    save_path = os.path.join(output_folder, filename)
    doc.save(save_path)
    return save_path

def download_link(path):
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    name = os.path.basename(path)
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{name}">📥 Download: {name}</a>'
    st.markdown(href, unsafe_allow_html=True)

# === UI ===
st.title("📄 Railway Letter Generator")
letter_type = st.selectbox("📌 Select Letter Type", list(template_files.keys()))

# === Employee Data ===
emp_data = pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name=None)
sheet = st.selectbox("📋 Select Sheet", list(emp_data.keys()))
df = emp_data[sheet]
df["Display"] = df.apply(lambda row: f"{row[1]} - {row[2]} - {row[4]} - {row[5]}", axis=1)
emp_display = st.selectbox("👤 Select Employee", df["Display"].dropna())
emp_row = df[df["Display"] == emp_display].iloc[0]

# === Common Fields ===
pf = emp_row[1]
hrms = emp_row[2]
unit_full = str(emp_row[4])
unit = unit_full[:2] if len(unit_full) >= 2 else unit_full
station = emp_row[8]
ename_en = emp_row[5]
ename_hi = emp_row[13]
desig = emp_row[18]
short = emp_row[14]
letter_no = f"{short}/{unit}/{station}"
today = date.today()

# === Context Init ===
context = {
    "PFNumber": pf,
    "UnitNumber": unit_full,
    "Unit": unit,
    "EmployeeName": ename_hi,
    "Designation": desig,
    "ShortName": short,
    "LetterNo": letter_no,
}

# === Letter Type Specific Blocks ===
if letter_type == "Duty Letter (For Absent)":
    duty_mode = st.selectbox("Duty Mode", ["SF-11 & Duty Letter For Absent", "Duty Letter For Absent"])
    from_date = st.date_input("From Date")
    to_date = st.date_input("To Date", value=today)
    join_date = st.date_input("Join Date", value=to_date + timedelta(days=1))
    letter_date = st.date_input("Letter Date", value=today)

    days = (to_date - from_date).days + 1
    memo = f"आप बिना किसी पूर्व सूचना के दिनांक {from_date.strftime('%d-%m-%Y')} से {to_date.strftime('%d-%m-%Y')} तक कुल {days} दिवस कार्य से अनुपस्थित थे, जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है। अतः आप कामों व भूलो के फेहरिस्त धारा 1, 2 एवं 3 के उल्लंघन के दोषी पाए जाते है।"

    context.update({
        "FromDate": from_date.strftime("%d-%m-%Y"),
        "ToDate": to_date.strftime("%d-%m-%Y"),
        "JoinDate": join_date.strftime("%d-%m-%Y"),
        "DutyDate": join_date.strftime("%d-%m-%Y"),
        "LetterDate": letter_date.strftime("%d-%m-%Y"),
        "Memo": memo
    })

elif letter_type == "SF-11 For Other Reason":
    letter_date = st.date_input("Letter Date", value=today)
    memo_user = st.text_area("📌 Memorandum Text")
    memo = memo_user + " जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है।"
    context.update({
        "LetterDate": letter_date.strftime("%d-%m-%Y"),
        "Memo": memo
    })

elif letter_type == "Sick Memo":
    letter_date = st.date_input("Letter Date", value=today)
    context.update({"LetterDate": letter_date.strftime("%d-%m-%Y")})

elif letter_type == "General Letter":
    letter_date = st.date_input("Letter Date", value=today)
    subject = st.text_input("Subject")
    reference = st.text_input("Reference (Optional)")
    memo = st.text_area("Memo")
    copy_to = st.text_area("Copy To (Optional)")
    context.update({
        "LetterDate": letter_date.strftime("%d-%m-%Y"),
        "Subject": subject,
        "Reference": reference if reference else "",
        "Memo": memo,
        "CopyTo": copy_to if copy_to else ""
    })

elif letter_type == "Exam NOC":
    year = st.selectbox("NOC Year", [str(y) for y in range(today.year - 1, today.year + 2)])
    exam = st.text_input("Exam Name")
    letter_date = st.date_input("Letter Date", value=today)

    past = pd.read_excel(exam_noc_path)
    emp_noc = past[(past["PFNumber"] == pf) & (past["Year"] == int(year))]
    count = len(emp_noc)
    if count >= 4:
        st.error("❌ Already 4 NOCs issued this year!")
        st.stop()
    number = count + 1

    context.update({
        "LetterDate": letter_date.strftime("%d-%m-%Y"),
        "ExamName": exam,
        "NOCNumber": str(number),
        "NOCYear": year
    })

elif letter_type == "SF-11 Punishment Order":
    letter_date = st.date_input("Letter Date", value=today)
    memo = st.text_area("📌 Punishment Memo")
    context.update({
        "LetterDate": letter_date.strftime("%d-%m-%Y"),
        "Memo": memo
    })

# === Generate Button ===
if st.button("📄 Generate Letter"):
    template = template_files[letter_type]
    filename = f"{letter_type} - {ename_hi}.docx"
    path = generate_doc(template, context, filename)
    st.success(f"✅ {letter_type} generated successfully!")
    download_link(path)

    # SF-11 Register Entry
    if letter_type == "SF-11 For Other Reason":
        sf_data = pd.read_excel(sf11_register_path, sheet_name="SSE-SGAM")
        new_entry = pd.DataFrame([[pf, ename_hi, desig, letter_no, context["LetterDate"], memo]], columns=sf_data.columns[:6])
        sf_data = pd.concat([sf_data, new_entry], ignore_index=True)
        with pd.ExcelWriter(sf11_register_path, mode="a", if_sheet_exists="replace") as writer:
            sf_data.to_excel(writer, sheet_name="SSE-SGAM", index=False)

    # Exam NOC Register Entry
    if letter_type == "Exam NOC":
        new_noc = pd.DataFrame([{
            "PFNumber": pf,
            "EmployeeName": ename_hi,
            "Year": int(year),
            "ExamName": exam,
            "NOCNumber": number,
            "Date": context["LetterDate"]
        }])
        full_noc = pd.concat([past, new_noc], ignore_index=True)
        full_noc.to_excel(exam_noc_path, index=False)