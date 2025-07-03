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

    selected_sf11_row = sf11_register[sf11_register["Display"] == selected_sf11_display].iloc[0]

    # Extract master data for selected employee (PF No. match)
    pf_number = selected_sf11_row["पी.एफ. क्रमांक"]
    letter_no = selected_sf11_row["पत्र क्र."]
    hindi_name = selected_sf11_row["कर्मचारी का नाम"]

    # Find employee's designation from Master Data
    found = False
    for sheet in employee_master.values():
        match_row = sheet[sheet.iloc[:, 1] == pf_number]
        if not match_row.empty:
            designation = match_row.iloc[0][18]  # Column 18 = Designation (Hindi)
            short_name = match_row.iloc[0][14] if len(match_row.columns) > 14 else ""
            found = True
            break

    if not found:
        st.error("❌ Employee designation not found in master data.")
    else:
        # दण्ड आदेश क्रमांक (D-1 fixed + पत्र क्र.)
        order_no = f"D-1 / {letter_no}"

        # Memo taken from SF-11 row or fixed
        memo_text = selected_sf11_row["आरोप का विवरण"]
        letter_date = st.date_input("📅 Letter Date", date.today())

        # Fill context
        context = {
            "EmployeeName": f"{hindi_name}, {designation}",
            "LetterNo.": letter_no,
            "LetterDate": letter_date.strftime("%d-%m-%Y"),
            "MEMO": memo_text
        }

        # Generate letter
        if st.button("📄 Generate SF-11 Punishment Order"):
            doc = Document(template_files["SF-11 Punishment Order"])
            replace_placeholders(doc, context)

            file_name = f"SF11_Punishment_{hindi_name}_{letter_date.strftime('%d-%m-%Y')}"
            output_docx = os.path.join("/tmp", f"{file_name}.docx")
            doc.save(output_docx)

            st.success("✅ SF-11 Punishment Order Generated!")
            download_button(output_docx, f"⬇️ Download {file_name}.docx")

            pdf_path = convert_to_pdf(output_docx)
            if pdf_path and os.path.exists(pdf_path):
                st.success("📄 PDF also generated!")
                download_button(pdf_path, f"⬇️ Download {os.path.basename(pdf_path)}")
            else:
                st.warning("⚠️ PDF conversion failed or not supported.")
