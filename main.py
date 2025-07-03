# letter_generator_app.py
import streamlit as st
import pandas as pd
from datetime import date, timedelta
from docx import Document
from docx2pdf import convert
import os
import base64

# Template mapping
template_files = {
    "SF-11 Punishment Order": "assets/SF-11 Punishment Order Template.docx",
    "SF-11 For Other Reason": "assets/SF-11 Other Reason Template.docx",
    "Duty Letter (For Absent)": "assets/Absent Duty letter temp.docx",
    "Sick Memo": "assets/SICK MEMO temp.docx",
    "Exam NOC": "assets/Exam NOC Letter temp.docx",
    "General Letter": "assets/General Letter temp.docx"
}

# Load employee data
@st.cache_data
def load_employee_data():
    return pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name="Apr.25")

@st.cache_data
def load_sf11_register():
    return pd.read_excel("assets/SF-11 Register.xlsx", sheet_name="SSE-SGAM")

@st.cache_data
def load_exam_noc_report():
    return pd.read_excel("assets/ExamNOC_Report.xlsx")

# Replace placeholders
def replace_placeholders(doc, context):
    for p in doc.paragraphs:
        for key, val in context.items():
            if f"[{key}]" in p.text:
                p.text = p.text.replace(f"[{key}]", str(val))
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for key, val in context.items():
                    if f"[{key}]" in cell.text:
                        cell.text = cell.text.replace(f"[{key}]", str(val))

# Generate Word
def generate_docx(template_path, context, filename):
    doc = Document(template_path)
    replace_placeholders(doc, context)
    docx_path = os.path.join("/tmp", filename + ".docx")
    doc.save(docx_path)
    return docx_path

# Convert to PDF
def convert_to_pdf(docx_path):
    try:
        pdf_path = docx_path.replace(".docx", ".pdf")
        convert(docx_path, pdf_path)
        return pdf_path
    except:
        return None

# Download button
def download_button(file_path, label):
    with open(file_path, "rb") as f:
        data = f.read()
        b64 = base64.b64encode(data).decode()
        href = f'<a href="data:application/octet-stream;base64,{b64}" download="{os.path.basename(file_path)}">{label}</a>'
        st.markdown(href, unsafe_allow_html=True)

# ============ MAIN UI START ============
st.title("📄 Letter Generator")
selected_type = st.selectbox("✉️ Select Letter Type", list(template_files.keys()))

if selected_type == "SF-11 For Other Reason":
    df = load_employee_data()
    df["DisplayName"] = df.apply(lambda row: f"{row[1]} - {row[2]} - {str(row[4])[:2]}/{row[9]} - {row[5]}", axis=1)
    selected_emp_display = st.selectbox("👤 Select Employee", df["DisplayName"].dropna().tolist())

    if selected_emp_display:
        selected_row = df[df["DisplayName"] == selected_emp_display].iloc[0]
        english_name = selected_row[5]
        hindi_name = selected_row[13]
        pf_number = selected_row[1]
        designation = selected_row[18]
        shortname = selected_row[15]
        unit_short = str(selected_row[4])[:2]
        working_station = selected_row[9]
        letter_no = f"{shortname} / {unit_short}/{working_station}"

        memo = st.text_area("📝 Memorandum")
        final_memo = f"{memo}, जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है। अतः आप कामों व भूलो के फेहरिस्त धारा 1, 2 एवं 3 के उल्लंघन के दोषी पाए जाते है।"
        letter_date = st.date_input("📅 Letter Date", date.today())

        if st.button("📄 Generate SF-11 Letter"):
            context = {
                "LetterDate": letter_date.strftime("%d-%m-%Y"),
                "EmployeeName": hindi_name,
                "Designation": designation,
                "Memo": final_memo,
                "PFNumber": pf_number,
                "LetterNo": letter_no
            }
            filename = f"SF11_OtherReason_{english_name}_{letter_date.strftime('%d-%m-%Y')}"
            docx_path = generate_docx(template_files[selected_type], context, filename)
            st.success("✅ Word Letter generated!")
            download_button(docx_path, f"⬇️ Download {os.path.basename(docx_path)}")

            # Save to SF-11 Register
            try:
                sf_df = load_sf11_register()
                last_index = sf_df.shape[0] + 1
                new_entry = {
                    "स.क्र.": last_index,
                    "पी.एफ. क्रमांक": pf_number,
                    "कर्मचारी का नाम": hindi_name,
                    "पदनाम": designation,
                    "पत्र क्र.": letter_no,
                    "दिनांक": letter_date.strftime("%d-%m-%Y"),
                    "आरोप का विवरण": final_memo
                }
                updated_df = pd.concat([sf_df, pd.DataFrame([new_entry])], ignore_index=True)
                updated_df.to_excel("assets/SF-11 Register.xlsx", sheet_name="SSE-SGAM", index=False)
                st.success("📘 Entry added to SF-11 Register")
            except Exception as e:
                st.error(f"❌ Error updating register: {e}")

            pdf_path = convert_to_pdf(docx_path)
            if pdf_path and os.path.exists(pdf_path):
                st.success("📄 PDF generated!")
                download_button(pdf_path, f"⬇️ Download {os.path.basename(pdf_path)}")
            else:
                st.warning("⚠️ PDF conversion failed or not supported.")

# You can extend further for Duty Letter, Exam NOC, General Letter similarly.