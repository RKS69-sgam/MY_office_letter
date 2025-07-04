import streamlit as st
import pandas as pd
from datetime import date, timedelta
from docx import Document
from tempfile import NamedTemporaryFile
import base64
import os
#from docx2pdf import convert

# Template Mapping
template_files = {
    "Duty Letter (For Absent)": "assets/Absent Duty letter temp.docx",
    "SF-11 For Other Reason": "assets/SF-11 temp.docx",
    "Sick Memo": "assets/SICK MEMO temp..docx",
    "General Letter": "assets/General Letter temp.docx",
    "Exam NOC": "assets/Exam NOC Letter temp.docx",
    "SF-11 Punishment Order": "assets/SF-11 Punishment order temp.docx"
}

# Dropdown for letter type
letter_types = [
    "Duty Letter (For Absent)",
    "SF-11 For Other Reason",
    "Sick Memo",
    "General Letter",
    "Exam NOC",
    "SF-11 Punishment Order"
]
selected_letter_type = st.selectbox("📌 Select Letter Type:", letter_types)

# Load Employee Master Data
employee_master = pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name=None)
sheet_names = list(employee_master.keys())
selected_sheet = st.selectbox("📋 Select Sheet", sheet_names)
df_emp = employee_master[selected_sheet]
df_emp["Display"] = df_emp.apply(lambda row: f"{row[1]} - {row[2]} - {row[4]} - {row[5]}", axis=1)
emp_display_list = df_emp["Display"].dropna().tolist()
selected_emp_display = st.selectbox("👤 Select Employee:", emp_display_list)
selected_row = df_emp[df_emp["Display"] == selected_emp_display].iloc[0]

# === DUTY LETTER SECTION ===
if selected_letter_type == "Duty Letter (For Absent)":
    st.subheader("📄 Generate Duty Letter (For Absent)")

    duty_mode = st.selectbox("🛠 Select Duty Letter Mode", [
        "SF-11 & Duty Letter For Absent",
        "Duty Letter For Absent"
    ])

    from_date = st.date_input("📅 From Date")
    to_date = st.date_input("📅 To Date", value=date.today())
    join_date = st.date_input("📆 Join Date", value=to_date + timedelta(days=1))
    letter_date = st.date_input("📄 Letter Date", value=date.today())

    # Get Employee Info
    pf_number = selected_row[1]
    hrms_id = selected_row[2]
    unit_raw = selected_row[4]
    working_station = selected_row[8]
    english_name = selected_row[5]
    hindi_name = selected_row[13]
    designation = selected_row[18]
    short_name = selected_row[14] if len(selected_row) > 14 else ""

    # Placeholder context
    context = {
        "LetterDate": letter_date.strftime("%d-%m-%Y"),
        "EmployeeName": hindi_name,
        "Designation": designation,
        "FromDate": from_date.strftime("%d-%m-%Y"),
        "ToDate": to_date.strftime("%d-%m-%Y"),
        "JoinDate": join_date.strftime("%d-%m-%Y"),
        "PFNumber": pf_number
    }

    def generate_doc(template_path, context):
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
        temp_file = NamedTemporaryFile(delete=False, suffix=".docx")
        doc.save(temp_file.name)
        return temp_file.name

    def download_file(file_path):
        with open(file_path, "rb") as f:
            data = f.read()
        b64 = base64.b64encode(data).decode()
        href = f'<a href="data:file/docx;base64,{b64}" download="{os.path.basename(file_path)}">📥 Download {os.path.basename(file_path)}</a>'
        st.markdown(href, unsafe_allow_html=True)

    if st.button("📄 Generate Duty Letter"):
        filled_doc = generate_doc(template_files["Duty Letter (For Absent)"], context)

        st.success("✅ Duty Letter generated successfully!")
        download_file(filled_doc)

        # Try converting to PDF
        try:
            pdf_path = filled_doc.replace(".docx", ".pdf")
            convert(filled_doc, pdf_path)
            if os.path.exists(pdf_path):
                st.success("📄 PDF also generated!")
                download_file(pdf_path)
            else:
                st.warning("⚠️ PDF file not found after conversion.")
        except:
            st.warning("⚠️ PDF conversion failed or not supported.")