import streamlit as st
import pandas as pd
from datetime import date
from docx import Document
import base64
import os
from tempfile import NamedTemporaryFile

# === Template Files ===
template_files = {
    "SF-11 Punishment Order": "assets/SF-11 Punishment order temp.docx",
    "SF-11 For Other Reason": "assets/SF-11 temp.docx",
    "Duty Letter (For Absent)": "assets/Absent Duty letter temp.docx",
    "Sick Memo": "assets/SICK MEMO temp..docx",
    "Exam NOC": "assets/Exam NOC Letter temp.docx",
    "General Letter": "assets/General Letter temp.docx"
}

# === Reload Button ===
if st.button("🔁 Reload All Data"):
    st.cache_data.clear()
    st.experimental_rerun()

# === Cache-Free Excel Loaders ===
@st.cache_data(ttl=0)
def load_employee_master():
    return pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name=None)

@st.cache_data(ttl=0)
def load_sf11_register():
    return pd.read_excel("assets/SF-11 Register.xlsx", sheet_name="SSE-SGAM")

@st.cache_data(ttl=0)
def load_exam_noc_data():
    return pd.read_excel("assets/ExamNOC_Report.xlsx")

# === Document Functions ===
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

# === Main UI ===
st.title("📄 Letter Generator")
st.success("Template files mapped correctly and ready to use.")

# Example Dropdown with Reload
letter_types = list(template_files.keys())
selected_letter_type = st.selectbox("1️⃣ Select Letter Type:", letter_types, key="letter_type")

# Load fresh data on every run
employee_master = load_employee_master()
sf11_register = load_sf11_register()
exam_noc_data = load_exam_noc_data()

# You can now add logic here using these fresh data
# SF-11 Punishment Order Section
if selected_letter_type == "SF-11 Punishment Order":
    st.subheader("📄 SF-11 Punishment Order Letter")

    # Dropdown to select employee from SF-11 Register
    sf11_register["Display"] = sf11_register.apply(
        lambda row: f"{row['पी.एफ. क्रमांक']} - {row['कर्मचारी का नाम']} - {row['दिनांक']} - {row['पत्र क्र.']}",
        axis=1
    )
    sf11_display_list = sf11_register["Display"].tolist()
    selected_display = st.selectbox("👤 Select Employee (SF-11 Register):", sf11_display_list)

    # Extract selected row
    selected_row = sf11_register[sf11_register["Display"] == selected_display].iloc[0]

    # Extract values
    pf_number = selected_row["पी.एफ. क्रमांक"]
    hindi_name = selected_row["कर्मचारी का नाम"]
    designation = selected_row["पदनाम"]
    letter_no = selected_row["पत्र क्र."]
    letter_date = st.date_input("📅 Letter Date", date.today())
    memo = selected_row["आरोप का विवरण"]

    # Generate D-1 format number
    final_letter_no = f"11 / D-1"

    # Generate context
    context = {
        "EmployeeName": f"{hindi_name} {designation}",
        "LetterNo.": letter_no,
        "LetterDate": letter_date.strftime("%d-%m-%Y"),
        "MEMO": memo
    }

    if st.button("📄 Generate SF-11 Punishment Order"):
        doc_path = generate_doc(template_files["SF-11 Punishment Order"], context)
        st.success("✅ Letter generated successfully!")

        download_file(doc_path)

        pdf_path = convert_to_pdf(doc_path)
        if pdf_path and os.path.exists(pdf_path):
            st.success("📄 PDF also ready!")
            download_file(pdf_path)
        else:
            st.warning("⚠️ PDF not generated.")