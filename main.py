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

    sf11_register["Display"] = sf11_register.apply(
        lambda row: f"{row['पी.एफ. क्रमांक']} - {row['कर्मचारी का नाम']} - {row['दिनांक']} - {row['पत्र क्र.']}", axis=1
    )
    sf11_employee_list = sf11_register["Display"].tolist()
    selected_sf11_display = st.selectbox("👤 Select Employee (SF-11 Register):", sf11_employee_list)

    if selected_sf11_display:
        selected_row = sf11_register[sf11_register["Display"] == selected_sf11_display].iloc[0]
        pf_number = selected_row["पी.एफ. क्रमांक"]
        hindi_name = selected_row["कर्मचारी का नाम"]
        letter_no = selected_row["पत्र क्र."]
        designation = selected_row["पदनाम"] if "पदनाम" in selected_row else ""

        letter_date = st.date_input("📅 Letter Date", date.today())

        reply_status = st.selectbox("📨 प्रति उत्तर प्राप्त हुआ?", ["हाँ", "नहीं"])
        punishment_number = f"D-1/{letter_no}"

        st.markdown(f"### 🔢 दण्डादेश क्रमांक: `{punishment_number}`")

        punishment_options = [
            "आगामी देय एक वर्ष की वेतन वृद्धि असंचयी प्रभाव से रोके जाने के अर्थदंड से दंडित किया जाता है।",
            "आगामी देय एक वर्ष की वेतन वृद्धि संचयी प्रभाव से रोके जाने के अर्थदंड से दंडित किया जाता है।",
            "आगामी देय एक सेट सुविधा पास तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।",
            "आगामी देय एक सेट PTO तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।",
            "आगामी देय दो सेट सुविधा पास तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।",
            "आगामी देय दो सेट PTO तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।"
        ]
        selected_punishment = st.selectbox("⚖️ दण्ड का विवरण:", punishment_options)

        appeal_date = st.date_input("📅 अपील दिनांक (यदि हो)", value=None)
        appeal_detail = st.text_area("📝 अपील निर्णय पत्र क्र. एवं विवरण", "")
        remark = st.text_area("📌 रिमार्क", "")

        if st.button("📄 Generate Punishment Order"):
            context = {
                "LetterDate": letter_date.strftime("%d-%m-%Y"),
                "Name": hindi_name,
                "Designation": designation,
                "LetterNo": letter_no,
                "PunishmentOrderNo": punishment_number,
                "PunishmentDetail": selected_punishment,
                "AppealDate": appeal_date.strftime("%d-%m-%Y") if appeal_date else "",
                "AppealDetail": appeal_detail,
                "Remark": remark,
            }

            output_path = generate_doc(template_files["SF-11 Punishment Order"], context)
            st.success("✅ Punishment Order Generated Successfully!")
            download_button(output_path, f"Download_Punishment_Order_{hindi_name}.docx")

            pdf_path = convert_to_pdf(output_path)
            if pdf_path and os.path.exists(pdf_path):
                st.success("📄 PDF also generated!")
                download_button(pdf_path, f"Download_Punishment_Order_{hindi_name}.pdf")
            else:
                st.warning("⚠️ PDF conversion failed or not supported.")

            # Save entry to SF-11 Register
            try:
                sf11_df = pd.read_excel("assets/SF-11 Register.xlsx", sheet_name="SSE-SGAM")
                new_entry = {
                    "स.क्र.": len(sf11_df) + 1,
                    "पी.एफ. क्रमांक": pf_number,
                    "कर्मचारी का नाम": hindi_name,
                    "पदनाम": designation,
                    "पत्र क्र.": letter_no,
                    "दिनांक": letter_date.strftime("%d-%m-%Y"),
                    "आरोप का विवरण": selected_punishment,
                    "दण्डादेश क्रमांक": punishment_number,
                    "दण्डादेश जारी करने का दिनांक": letter_date.strftime("%d-%m-%Y"),
                    "यदि कर्मचारी द्वारा अपील किया गया हो तो अपील का दिनांक": appeal_date.strftime("%d-%m-%Y") if appeal_date else "",
                    "अपील निर्णय पत्र क्र. एवं संक्षिप्त विवरण": appeal_detail,
                    "रिमार्क": remark,
                    "प्रत्युत्तर प्राप्‍त हुआ": reply_status
                }
                updated_df = pd.concat([sf11_df, pd.DataFrame([new_entry])], ignore_index=True)
                updated_df.to_excel("assets/SF-11 Register.xlsx", sheet_name="SSE-SGAM", index=False)
                st.success("🗂 Entry Saved to SF-11 Register.")
            except Exception as e:
                st.error(f"❌ Error saving to register: {e}")
