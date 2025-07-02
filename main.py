import streamlit as st
import pandas as pd
from datetime import date, timedelta
from docx import Document
import base64
import os

# === Constants ===
EMPLOYEE_FILE = "assets/EMPLOYEE MASTER DATA.xlsx"
SF11_FILE = "assets/SF-11 Register.xlsm"
TEMPLATE_FILES = {
    "SF-11 Punishment Order": "assets/SF-11 temp.docx",
    "Duty Letter (For Absent)": "assets/Absent Duty letter temp.docx",
    "Sick Memo": "assets/SICK MEMO temp..docx",
    "Exam NOC": "assets/Exam NOC Letter temp.docx",
    "General Letter": "assets/General Letter temp.docx"
}

# === Helper functions ===
def load_excel_data(file_path, sheet_name=None):
    return pd.read_excel(file_path, sheet_name=sheet_name)

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

def generate_docx(template_path, context, filename):
    doc = Document(template_path)
    replace_placeholders(doc, context)
    save_path = os.path.join("/tmp", filename + ".docx")
    doc.save(save_path)
    return save_path

def download_button(file_path, label):
    with open(file_path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
        href = f'<a href="data:application/octet-stream;base64,{b64}" download="{os.path.basename(file_path)}">{label}</a>'
        st.markdown(href, unsafe_allow_html=True)

# === Top-Level: Letter Type Selection ===
letter_type = st.selectbox("📂 Select Letter Type:", [
    "SF-11 Punishment Order",
    "Duty Letter (For Absent)",
    "Sick Memo",
    "Exam NOC",
    "General Letter"
])

# === Load Data ===
if letter_type == "SF-11 Punishment Order":
    df = load_excel_data(SF11_FILE, sheet_name="SSE-SGAM")
    df = df.dropna(subset=["Employee Name", "पत्र क्रमांक"])
    df["Dropdown"] = df.apply(lambda row: f"{row['Employee Name']} ({row['पत्र क्रमांक']})", axis=1)
    selected_emp = st.selectbox("Select SF-11 Entry:", df["Dropdown"].tolist())
    row = df[df["Dropdown"] == selected_emp].iloc[0]

    pf_no = row["PF No."]
    emp_name = row["Employee Name"]
    unit = row["Unit"]
    designation = row["Designation"]
    letter_no = row["पत्र क्रमांक"]
    memo_text = row["Memo"]
    letter_date = st.date_input("Letter Date", date.today())

    context = {
        "PFNumber": pf_no,
        "EmployeeName": emp_name,
        "UnitNumber": unit,
        "Designation": designation,
        "LetterDate": letter_date.strftime("%d-%m-%Y"),
        "MEMO": memo_text,
        "LetterNo": letter_no
    }

else:
    # Load Employee Master
    emp_data = pd.read_excel(EMPLOYEE_FILE, sheet_name=None)
    sheet_names = list(emp_data.keys())
    selected_sheet = st.selectbox("Select Unit Sheet:", sheet_names)
    df = emp_data[selected_sheet]

    df["DisplayUnit"] = df.apply(lambda row: f"{str(row[4])[:2]}/{row[8]}", axis=1)
    df["DisplayName"] = df.apply(
        lambda row: f"PF:{row[1]}, HRMS:{row[2]}, Unit:{row['DisplayUnit']}, {row[5]}", axis=1)
    selected_emp = st.selectbox("Select Employee:", df["DisplayName"].dropna().tolist())
    row = df[df["DisplayName"] == selected_emp].iloc[0]

    emp_name = row[13]   # Hindi name
    pf_no = row[1]
    designation = row[18]
    unit = f"{str(row[4])[:2]}/{row[8]}"
    letter_date = st.date_input("Letter Date", date.today())
    from_date = st.date_input("From Date") if "Duty" in letter_type else None
    to_date = st.date_input("To Date") if "Duty" in letter_type else None
    join_date = st.date_input("Join Duty Date", (to_date + timedelta(days=1)) if to_date else date.today()) if "Duty" in letter_type else None
    memo_text = st.text_area("Memo Text") if "Sick" in letter_type else ""
    exam_name = st.text_input("Exam Name") if "NOC" in letter_type else ""
    noc_count = st.selectbox("NOC Attempt No", [1,2,3,4]) if "NOC" in letter_type else None
    general_subject = st.text_input("Subject") if "General" in letter_type else ""
    general_body = st.text_area("Body") if "General" in letter_type else ""

    context = {
        "EmployeeName": emp_name,
        "PFNumber": pf_no,
        "Designation": designation,
        "UnitNumber": unit,
        "LetterDate": letter_date.strftime("%d-%m-%Y"),
        "FromDate": from_date.strftime("%d-%m-%Y") if from_date else "",
        "ToDate": to_date.strftime("%d-%m-%Y") if to_date else "",
        "DutyDate": join_date.strftime("%d-%m-%Y") if join_date else "",
        "MEMO": memo_text,
        "ExamName": exam_name,
        "NOCCount": noc_count,
        "Subject": general_subject,
        "Body": general_body
    }

# === Generate Letter ===
if st.button("Generate Letter"):
    filename = f"{letter_type.split()[0]}_{emp_name}_{letter_date.strftime('%d%m%Y')}"
    docx_path = generate_docx(TEMPLATE_FILES[letter_type], context, filename)
    st.success("Word Letter Generated ✅")
    download_button(docx_path, "⬇️ Download Word Letter")