import streamlit as st
import pandas as pd
from docx import Document
import os
from datetime import datetime
from docx2pdf import convert
import shutil

# Paths
sf11_register_path = "SF-11 Register.xlsx"
master_data_path = "EMPLOYEE MASTER DATA.xlsx"
template_path = "assets/sf11_punishment_template.docx"  # Your correct punishment template
output_folder = "generated_letters"
os.makedirs(output_folder, exist_ok=True)

# Load SF-11 Register
@st.cache_data
def load_sf11_register():
    return pd.read_excel(sf11_register_path, sheet_name="SSE-SGAM")

sf11_register = load_sf11_register()

# Display Column Creation (PF - Name - Date)
sf11_register["Display"] = (
    sf11_register["PF No."].astype(str)
    + " - "
    + sf11_register["कर्मचारी का नाम (हिंदी)"]
    + " - "
    + pd.to_datetime(sf11_register["Letter Date"]).dt.strftime("%d.%m.%Y")
)

st.title("📄 SF-11 Punishment Order Letter")

# Dropdown to select employee
selected_display = st.selectbox("👤 Select Employee (SF-11 Register):", sf11_register["Display"])
selected_row = sf11_register[sf11_register["Display"] == selected_display].iloc[0]

# Date picker
letter_date = st.date_input("📅 Letter Date", datetime.today()).strftime("%d-%m-%Y")

# Generate button
if st.button("📝 Generate SF-11 Punishment Order"):
    try:
        # Load Template
        doc = Document(template_path)

        # Fill Placeholders
        replacements = {
            "[EmployeeName]": selected_row["कर्मचारी का नाम (हिंदी)"] + " " + selected_row["पदनाम (हिंदी)"],
            "[LetterNo.]": selected_row["पत्र क्र."],
            "[LetterDate]": letter_date,
            "[MEMO]": selected_row["अनुशासनात्मक विवरण"],
        }

        for para in doc.paragraphs:
            for key, value in replacements.items():
                if key in para.text:
                    para.text = para.text.replace(key, value)

        # Save DOCX
        emp_name = selected_row["कर्मचारी का नाम (हिंदी)"]
        filename_docx = f"SF-11 - {emp_name} - {letter_date}.docx"
        filename_pdf = filename_docx.replace(".docx", ".pdf")
        docx_path = os.path.join(output_folder, filename_docx)
        pdf_path = os.path.join(output_folder, filename_pdf)

        doc.save(docx_path)

        # Convert to PDF
        try:
            convert(docx_path, pdf_path)
            with open(pdf_path, "rb") as f:
                st.success("✅ Letter generated successfully!")
                st.download_button("⬇️ Download PDF", f, file_name=filename_pdf)
        except:
            with open(docx_path, "rb") as f:
                st.warning("⚠️ PDF conversion failed. Download Word file instead.")
                st.download_button("⬇️ Download Word", f, file_name=filename_docx)

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")