import streamlit as st import pandas as pd from datetime import date, timedelta from docx import Document from docx2pdf import convert from tempfile import NamedTemporaryFile import base64 import os

=== Load Excel Data ===

employee_master = pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name=None)

=== UI Setup ===

st.title("📄 Duty Letter Generator (For Absent)") st.markdown("---")

=== Select Sheet and Employee ===

sheet_names = list(employee_master.keys()) selected_sheet = st.selectbox("📄 Select Sheet:", sheet_names) df_emp = employee_master[selected_sheet] df_emp["Display"] = df_emp.apply(lambda row: f"{row[1]} - {row[2]} - {row[4]} - {row[5]}", axis=1) emp_display_list = df_emp["Display"].dropna().tolist() selected_emp_display = st.selectbox("👤 Select Employee:", emp_display_list)

=== Fetch Employee Info ===

selected_row = df_emp[df_emp["Display"] == selected_emp_display].iloc[0] pf_number = selected_row[1] hindi_name = selected_row[13] designation = selected_row[18] english_name = selected_row[5] short_name = selected_row[14] if len(selected_row) > 14 else ""

=== Duty Dates ===

from_date = st.date_input("📅 From Date") to_date = st.date_input("📅 To Date", date.today()) join_date = st.date_input("📆 Join Date", to_date + timedelta(days=1)) duty_mode = st.selectbox("📌 Duty Letter Type", ["SF-11 & Duty Letter For Absent", "Duty Letter For Absent"]) letter_date = st.date_input("📄 Letter Date", date.today())

=== Template Context ===

context = { "EmployeeName": hindi_name, "Designation": designation, "PFNumber": pf_number, "FromDate": from_date.strftime("%d-%m-%Y"), "ToDate": to_date.strftime("%d-%m-%Y"), "JoinDate": join_date.strftime("%d-%m-%Y"), "DutyMode": duty_mode, "LetterDate": letter_date.strftime("%d-%m-%Y") }

=== Replace Placeholders in Word Template ===

def replace_placeholders(doc, context): for p in doc.paragraphs: for key, val in context.items(): if f"[{key}]" in p.text: p.text = p.text.replace(f"[{key}]", str(val)) for table in doc.tables: for row in table.rows: for cell in row.cells: for key, val in context.items(): if f"[{key}]" in cell.text: cell.text = cell.text.replace(f"[{key}]", str(val))

=== Generate Document ===

def generate_docx(template_path, context, filename): doc = Document(template_path) replace_placeholders(doc, context) output_path = os.path.join("/tmp", f"{filename}.docx") doc.save(output_path) return output_path

=== Download Button ===

def download_button(file_path, label): with open(file_path, "rb") as f: b64 = base64.b64encode(f.read()).decode() href = f'<a href="data:application/octet-stream;base64,{b64}" download="{os.path.basename(file_path)}">{label}</a>' st.markdown(href, unsafe_allow_html=True)

=== Generate on Button Click ===

if st.button("📄 Generate Duty Letter"): template_path = "assets/Absent Duty letter temp.docx" filename = f"Duty_Letter_{english_name}_{from_date.strftime('%d-%m-%Y')}" docx_path = generate_docx(template_path, context, filename)

st.success("✅ Word Document Created")
download_button(docx_path, "⬇️ Download Word")

try:
    pdf_path = docx_path.replace(".docx", ".pdf")
    convert(docx_path, pdf_path)
    st.success("📄 PDF Generated")
    download_button(pdf_path, "⬇️ Download PDF")
except Exception as e:
    st.warning(f"⚠️ PDF conversion failed: {e}")

