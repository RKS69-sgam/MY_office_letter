import streamlit as st
import pandas as pd
from datetime import date, timedelta
from docx import Document
import base64
import os

@st.cache_data
def load_data():
    master = pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name=None)
    sf11 = pd.read_excel("assets/SF-11 Register.xlsm", sheet_name="SSE-SGAM")
    return master, sf11

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

def correct_unit_format(unit, station):
    try:
        unit_str = str(unit).strip()
        if "/" in unit_str:
            unit_str = unit_str.split("/")[0].strip()
        if unit_str.isdigit():
            unit_str = unit_str[:2]
        return f"{unit_str}/{str(station).strip()}"
    except:
        return ""

def download_button(file_path, label):
    with open(file_path, "rb") as f:
        data = f.read()
        b64 = base64.b64encode(data).decode()
        href = f'<a href="data:application/octet-stream;base64,{b64}" download="{os.path.basename(file_path)}">{label}</a>'
        st.markdown(href, unsafe_allow_html=True)

def generate_docx(template_path, context, filename):
    doc = Document(template_path)
    replace_placeholders(doc, context)
    out_path = os.path.join("/tmp", filename + ".docx")
    doc.save(out_path)
    return out_path

# Load data
master_data, sf11_data = load_data()

st.title("ðŸ“„ Letter Generator")

letter_type = st.selectbox("Select Letter Type:", [
    "SF-11 Punishment Order",
    "Duty Letter (For Absent)",
    "Sick Memo",
    "Exam NOC",
    "General Letter"
])

# Determine source based on letter type
if letter_type == "SF-11 Punishment Order":
    sf11_data["Dropdown"] = sf11_data.apply(lambda r: f"{r[2]} ({r[6]})", axis=1)
    selected_display = st.selectbox("Select Employee:", sf11_data["Dropdown"].tolist())
    selected_row = sf11_data[sf11_data["Dropdown"] == selected_display].iloc[0]
    emp_name = selected_row[2]
    pf = selected_row[1]
    designation = selected_row[5]
    memo = selected_row[6]
    shortname = selected_row[14]
    context = {
        "LetterDate": date.today().strftime("%d-%m-%Y"),
        "EmployeeName": emp_name,
        "Designation": designation,
        "MEMO": memo,
        "ShortName": shortname
    }

else:
    sheet_names = list(master_data.keys())
    selected_sheet = st.selectbox("Select Unit Sheet:", sheet_names)
    df = master_data[selected_sheet]
    df["DisplayUnit"] = df.apply(lambda r: correct_unit_format(r[4], r[8]), axis=1)
    df["DisplayName"] = df.apply(lambda r: f"{r[1]} - {r[2]} - {r['DisplayUnit']} - {r[5]}", axis=1)
    selected_display = st.selectbox("Select Employee:", df["DisplayName"].tolist())
    row = df[df["DisplayName"] == selected_display].iloc[0]

    emp_name = row[13]
    pf = row[1]
    designation = row[18]
    unit = row["DisplayUnit"]
    context = {
        "LetterDate": date.today().strftime("%d-%m-%Y"),
        "EmployeeName": emp_name,
        "Designation": designation,
        "UnitNumber": unit,
        "PFNumber": pf,
        "MEMO": st.text_area("Memo Text") if "Sick" in letter_type or "Duty" in letter_type else "",
        "ExamName": st.text_input("Exam Name") if "NOC" in letter_type else "",
        "Body": st.text_area("Letter Body") if letter_type == "General Letter" else "",
        "Subject": st.text_input("Subject") if letter_type == "General Letter" else "",
        "NOCCount": st.selectbox("NOC Attempt No", [1, 2, 3, 4]) if "NOC" in letter_type else ""
    }

# Template paths
template_map = {
    "SF-11 Punishment Order": "assets/SF-11 temp.docx",
    "Duty Letter (For Absent)": "assets/Absent Duty letter temp.docx",
    "Sick Memo": "assets/SICK MEMO temp..docx",
    "Exam NOC": "assets/Exam NOC Letter temp.docx",
    "General Letter": "assets/General Letter temp.docx"
}

# Generate Letter
if st.button("Generate Letter"):
    filename = f"{letter_type.split()[0]}_{emp_name}_{date.today().strftime('%d-%m-%Y')}"
    docx_path = generate_docx(template_map[letter_type], context, filename)
    st.success("Word file created successfully.")
    download_button(docx_path, f"â¬‡ï¸ Download {os.path.basename(docx_path)}")