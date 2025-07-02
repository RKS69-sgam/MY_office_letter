
import streamlit as st
import pandas as pd
from datetime import date, timedelta
from docx import Document
import base64
import os
import shutil
from tempfile import NamedTemporaryFile

# === Load data files ===
@st.cache_data
def load_employee_data():
    return pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name=None)

@st.cache_data
def load_sf11_register():
    return pd.read_excel("assets/SF-11 Register.xlsx", sheet_name="SSE-SGAM")

@st.cache_data
def load_exam_noc_data():
    return pd.read_excel("assets/ExamNOC_Report.xlsx")

@st.cache_data
def load_general_letter_data():
    return pd.read_excel("assets/General Letter.xlsx")

# === Replace placeholders ===
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

# === Generate filled Word file ===
def generate_docx(template_path, context, filename):
    doc = Document(template_path)
    replace_placeholders(doc, context)
    docx_path = os.path.join("/tmp", filename + ".docx")
    doc.save(docx_path)
    return docx_path

# === Download button ===
def download_button(file_path, label):
    with open(file_path, "rb") as f:
        data = f.read()
        b64 = base64.b64encode(data).decode()
        href = f'<a href="data:application/octet-stream;base64,{b64}" download="{os.path.basename(file_path)}">{label}</a>'
        st.markdown(href, unsafe_allow_html=True)

# === Main App ===
st.title("Letter Generator - Railway Office")

# Step 1: Letter Type
letter_type = st.selectbox("Select Letter Type", [
    "SF-11 Punishment Order",
    "SF-11 For Other Reason",
    "Duty Letter (For Absent)",
    "Sick Memo",
    "Exam NOC",
    "General Letter"
])

# Templates
template_files = {
    "SF-11 Punishment Order": "assets/SF-11 temp.docx",
    "SF-11 For Other Reason": "assets/SF-11 temp.docx",
    "Duty Letter (For Absent)": "assets/Absent Duty letter temp.docx",
    "Sick Memo": "assets/SICK MEMO temp..docx",
    "Exam NOC": "assets/Exam NOC Letter temp.docx",
    "General Letter": "assets/General Letter temp.docx"
}

# Placeholder implementation: full logic depends on the selected type (already developed)
st.markdown("ðŸ‘‰ Continue implementing UI and logic based on detailed conditional branches from the instructions.")