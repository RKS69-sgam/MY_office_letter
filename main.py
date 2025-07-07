# FINAL STREAMLIT LETTER GENERATOR APP WITH ALL LETTER TYPES FIXED

import streamlit as st
import pandas as pd
import os
import base64
from docx import Document
from docx.shared import Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import date, timedelta

# Create output folder
os.makedirs("generated_letters", exist_ok=True)

# File paths
template_files = {
    "Duty Letter (For Absent)": "assets/Absent Duty letter temp.docx",
    "SF-11 For Other Reason": "assets/SF-11 temp.docx",
    "Sick Memo": "assets/SICK MEMO temp..docx",
    "General Letter": "assets/General Letter temp.docx",
    "Exam NOC": "assets/Exam NOC Letter temp.docx",
    "SF-11 Punishment Order": "assets/SF-11 Punishment order temp.docx"
}

employee_master = pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name=None)
sf11_register_path = "assets/SF-11 Register.xlsx"
sf11_register = pd.read_excel(sf11_register_path, sheet_name="SSE-SGAM")
noc_register_path = "assets/Exam NOC_Report.xlsx"
df_noc = pd.read_excel(noc_register_path) if os.path.exists(noc_register_path) else pd.DataFrame(columns=["PF Number", "Employee Name", "Designation", "NOC Year", "Application No.", "Exam Name"])

# === Utility Functions ===
def replace_placeholder_in_para(paragraph, context):
    full_text = ''.join(run.text for run in paragraph.runs)
    new_text = full_text
    for key, val in context.items():
        new_text = new_text.replace(f"[{key}]", str(val))
    if new_text != full_text:
        for run in paragraph.runs:
            run.text = ''
        if paragraph.runs:
            paragraph.runs[0].text = new_text
        else:
            paragraph.add_run(new_text)

def set_table_border(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '6')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')
        borders.append(border)
    tblPr.append(borders)

def insert_exam_noc_table(doc, table_data):
    for i, p in enumerate(doc.paragraphs):
        if "[PFNumber]" in p.text:
            p.text = ""
            table = doc.tables[0]._parent.add_table(rows=0, cols=len(table_data[0]))
            table.autofit = True
            hdr_cells = table.add_row().cells
            for j, val in enumerate(table_data[0]):
                hdr_cells[j].text = str(val)
            for data_row in table_data[1:]:
                row_cells = table.add_row().cells
                for j, val in enumerate(data_row):
                    row_cells[j].text = str(val)
            set_table_border(table)
            return

def generate_word(template_path, context, filename, table_data=None):
    doc = Document(template_path)
    for p in doc.paragraphs:
        replace_placeholder_in_para(p, context)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_placeholder_in_para(p, context)
    if table_data:
        insert_exam_noc_table(doc, table_data)
    output_path = os.path.join("generated_letters", filename)
    doc.save(output_path)
    return output_path

def download_word(path):
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    name = os.path.basename(path)
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{name}">Download Word File</a>'
    st.markdown(href, unsafe_allow_html=True)

# === UI ===
st.title("OFFICE OF THE SSE/PW/SGAM")
letter_type = st.selectbox("Select Letter Type", list(template_files.keys()))

# === Employee Selection ===
if letter_type == "SF-11 Punishment Order":
    df = sf11_register
    df["Display"] = df.apply(lambda r: f"{r['पी.एफ. क्रमांक']} - {r['कर्मचारी का नाम']} - {r['पत्र क्र.']}", axis=1)
    selected = st.selectbox("Select Employee", df["Display"].dropna())
    row = df[df["Display"] == selected].iloc[0]
    patra_kr = row["पत्र क्र."]
    dandadesh_krmank = f"{patra_kr}/D-1"
    pf = row["पी.एफ. क्रमांक"]
    hname = row["कर्मचारी का नाम"]
    desg = row.get("पदनाम", "")
    unit_full = patra_kr.split("/")[1]
    unit = unit_full[:2]
    short = patra_kr.split("/")[0]
    letter_no = dandadesh_krmank
else:
    df = employee_master["Apr.25"]
    df["Display"] = df.apply(lambda r: f"{r[1]} - {r[2]} - {r[4]} - {r[5]}", axis=1)
    selected = st.selectbox("Select Employee", df["Display"].dropna())
    row = df[df["Display"] == selected].iloc[0]
    pf = row[1]
    hname = row[13]
    desg = row[18]
    unit_full = str(row[4])
    unit = unit_full[:2]
    short = row[14]
    letter_no = f"{short}/{unit}/{unit_full}"

letter_date = st.date_input("Letter Date", value=date.today())
context = {
    "LetterDate": letter_date.strftime("%d-%m-%Y"),
    "EmployeeName": hname,
    "Designation": desg,
    "PFNumber": pf,
    "ShortName": short,
    "Unit": unit,
    "UnitNumber": unit_full,
    "LetterNo": letter_no,
    "DutyDate": "",
    "FromDate": "",
    "ToDate": "",
    "JoinDate": "",
    "Memo": "",
    "OfficerUnit": "",
    "Subject": "",
    "Reference": "",
    "CopyTo": ""
}

if letter_type == "Duty Letter (For Absent)":
    fd = st.date_input("From Date")
    td = st.date_input("To Date")
    jd = st.date_input("Join Date", value=td + timedelta(days=1))
    context.update({
        "FromDate": fd.strftime("%d-%m-%Y"),
        "ToDate": td.strftime("%d-%m-%Y"),
        "JoinDate": jd.strftime("%d-%m-%Y"),
        "DutyDate": jd.strftime("%d-%m-%Y"),
        "Memo": f"आप बिना किसी पूर्व सूचना के दिनांक {fd.strftime('%d-%m-%Y')} से {td.strftime('%d-%m-%Y')} तक कुल {(td-fd).days+1} दिवस कार्य से अनुपस्थित थे, जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है। अतः आप कामों व भूलो के फेहरिस्त धारा 1, 2 एवं 3 के उल्लंघन के दोषी पाए जाते है।"
    })
elif letter_type == "SF-11 For Other Reason":
    memo_input = st.text_area("Memo")
    context["Memo"] = memo_input + " जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है। अतः आप कामों व भूलो के फेहरिस्त धारा 1, 2 एवं 3 के उल्लंघन के दोषी पाए जाते है।"
elif letter_type == "General Letter":
    context["OfficerUnit"] = st.text_input("Officer/Unit")
    context["Subject"] = st.text_input("Subject")
    context["Reference"] = st.text_input("Reference")
    context["Memo"] = st.text_area("Memo")
    context["CopyTo"] = st.text_input("Copy To")
elif letter_type == "Exam NOC":
    exam_name = st.text_input("Exam Name")
    year = st.selectbox("NOC Year", [2025, 2024])
    count = sum((df_noc["PF Number"] == pf) & (df_noc["NOC Year"] == year))
    if count >= 4:
        st.warning("Already 4 NOCs taken.")
    else:
        application_no = count + 1
        table_data = [["PF Number", "Employee Name", "Designation", "NOC Year", "Application No.", "Exam Name"],
                      [pf, hname, desg, year, application_no, exam_name]]
if st.button("Generate Letter"):
    file = template_files[letter_type]
    name = f"{letter_type.replace('/', '-')}-{hname}.docx"
    if letter_type == "Exam NOC" and count < 4:
        word_path = generate_word(file, context, name, table_data)
        download_word(word_path)
        new_noc = {"PF Number": pf, "Employee Name": hname, "Designation": desg, "NOC Year": year, "Application No.": application_no, "Exam Name": exam_name}
        df_noc = pd.concat([df_noc, pd.DataFrame([new_noc])], ignore_index=True)
        df_noc.to_excel(noc_register_path, index=False)
    else:
        word_path = generate_word(file, context, name)
        download_word(word_path)
