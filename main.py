
import streamlit as st
import pandas as pd
from datetime import date, timedelta
from docx import Document
from tempfile import NamedTemporaryFile
import base64
import os
import shutil

# === Load data ===
@st.cache_data
def load_data():
    sf11_data = pd.read_excel("assets/SF-11 Register.xlsx", sheet_name=None)
    master_data = pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name=None)
    return sf11_data, master_data

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
    docx_path = os.path.join("/tmp", filename + ".docx")
    doc.save(docx_path)
    return docx_path

def convert_to_pdf(docx_path):
    try:
        from docx2pdf import convert
        pdf_path = docx_path.replace(".docx", ".pdf")
        convert(docx_path, pdf_path)
        return pdf_path
    except:
        return None

def download_button(file_path, label):
    with open(file_path, "rb") as f:
        data = f.read()
        b64 = base64.b64encode(data).decode()
        href = f'<a href="data:application/octet-stream;base64,{b64}" download="{os.path.basename(file_path)}">{label}</a>'
        st.markdown(href, unsafe_allow_html=True)

# === MAIN ===
sf11_data, master_data = load_data()

letter_type = st.selectbox("Select Letter Type:", [
    "SF-11 Punishment Order",
    "Duty Letter (For Absent)",
    "Sick Memo",
    "Exam NOC",
    "General Letter"
])

# Step 2: Choose unit sheet only for general letters
if letter_type != "SF-11 Punishment Order":
    unit_sheet = st.selectbox("Select Unit Sheet:", list(master_data.keys()))
    selected_sheet = master_data[unit_sheet]
    employee_options = []
    for i, row in selected_sheet.iterrows():
        pf = str(row[1]).strip()
        hrms = str(row[2]).strip()
        unit = str(row[3]).strip()
        name = str(row[5]).strip()
        employee_options.append(f"{pf} - {hrms} - {unit}/{name}")
else:
    selected_sheet = sf11_data["SSE-SGAM"]
    employee_options = []
    for i, row in selected_sheet.iterrows():
        pf = str(row[1]).strip()
        name = str(row[2]).strip()
        letter_no = str(row[5]).strip()
        letter_date = str(row[6]).strip()
        employee_options.append(f"{pf} - {letter_no} - {letter_date} - {name}")

selected_emp = st.selectbox("Select Employee:", employee_options)
letter_date = st.date_input("Select Letter Date", value=date.today())
memo_text = ""
if letter_type == "SF-11 Punishment Order":
    memo_text = st.text_area("Memo Text")

# Parse selected employee
emp_parts = selected_emp.split(" - ")
context = {}

if letter_type == "SF-11 Punishment Order":
    df = sf11_data["SSE-SGAM"]
    row = df[df[df.columns[1]].astype(str) == emp_parts[0]].iloc[0]
    context["EmployeeName"] = row[2]
    context["Memo"] = memo_text
    context["LetterNo"] = row[5]
    context["LetterDate"] = letter_date.strftime("%d-%m-%Y")
    context["ShortName"] = row[14] if len(row) >= 15 else ""
    template_path = "assets/SF-11 temp.docx"
    file_name = f"SF-11 - {row[2]}"
elif letter_type == "Duty Letter (For Absent)":
    df = selected_sheet
    row = df[df[df.columns[1]].astype(str) == emp_parts[0]].iloc[0]
    context["EmployeeName"] = row[13]
    context["Designation"] = row[18]
    context["JoinDate"] = (date.today() + timedelta(days=1)).strftime("%d-%m-%Y")
    template_path = "assets/Absent Duty letter temp.docx"
    file_name = f"DutyLetter - {row[13]}"
elif letter_type == "Sick Memo":
    df = selected_sheet
    row = df[df[df.columns[1]].astype(str) == emp_parts[0]].iloc[0]
    context["EmployeeName"] = row[13]
    context["Designation"] = row[18]
    context["LetterDate"] = letter_date.strftime("%d-%m-%Y")
    template_path = "assets/SICK MEMO temp.docx"
    file_name = f"SickMemo - {row[13]}"
elif letter_type == "Exam NOC":
    df = selected_sheet
    row = df[df[df.columns[1]].astype(str) == emp_parts[0]].iloc[0]
    context["EmployeeName"] = row[13]
    context["Designation"] = row[18]
    template_path = "assets/Exam NOC Letter temp.docx"
    file_name = f"ExamNOC - {row[13]}"
elif letter_type == "General Letter":
    df = selected_sheet
    row = df[df[df.columns[1]].astype(str) == emp_parts[0]].iloc[0]
    context["EmployeeName"] = row[13]
    context["Designation"] = row[18]
    template_path = "assets/General Letter temp.docx"
    file_name = f"GeneralLetter - {row[13]}"

if st.button("Generate Letter"):
    docx_path = generate_docx(template_path, context, file_name)
    pdf_path = convert_to_pdf(docx_path)
    st.success("Letter generated successfully!")
    download_button(docx_path, "ðŸ“„ Download Word File")
    if pdf_path:
        download_button(pdf_path, "ðŸ“„ Download PDF File")