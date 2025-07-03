import streamlit as st
import pandas as pd
from docx import Document
import os
from datetime import datetime

# Try importing docx2pdf for PDF conversion (works only on Windows)
try:
    from docx2pdf import convert
    DOCX2PDF_AVAILABLE = True
except ImportError:
    DOCX2PDF_AVAILABLE = False

# Define paths
sf11_register_path = "SF-11 Register.xlsx"
template_path = "assets/sf11_punishment_template.docx"
output_folder = "generated_letters"
os.makedirs(output_folder, exist_ok=True)

# Load SF-11 Register
def load_sf11_register():
    return pd.read_excel(sf11_register_path, sheet_name="SSE-SGAM")

sf11_register = load_sf11_register()

# Create a display column for dropdown
sf11_register["Display"] = (
    sf11_register["PF No."].astype(str)
    + " - "
    + sf11_register["कर्मचारी का नाम (हिंदी)"]
    + " - "
    + pd.to_datetime(sf11_register["Letter Date"]).dt.strftime("%d.%m.%Y")
)

# Streamlit UI
st.title("📄 SF-11 Punishment Order Letter")
selected_display = st.selectbox("👤 Select Employee (SF-11 Register):", sf11_register["Display"])
selected_row = sf11_register[sf11_register["Display"] == selected_display].iloc[0]
letter_date = st.date_input("🗕️ Letter Date", datetime.today()).strftime("%d-%m-%Y")

if st.button("📝 Generate SF-11 Punishment Order"):
    try:
        # Load template
        doc = Document(template_path)

        # Placeholder replacements
        replacements = {
            "[EmployeeName]": selected_row["\u0915\u0930\u094d\u092e\u091a\u093e\u0930\u0940 \u0915\u093e \u0928\u093e\u092e (\u0939\u093f\u0902\u0926\u0940)"] + " " + selected_row["\u092a\u0926\u0928\u093e\u092e (\u0939\u093f\u0902\u0926\u0940)"],
            "[LetterNo.]": selected_row["\u092a\u0924\u094d\u0930 \u0915\u094d\u0930."],
            "[LetterDate]": letter_date,
            "[MEMO]": selected_row["\u0905\u0928\u0941\u0936\u093e\u0938\u0928\u093e\u0924\u094d\u092e\u0915 \u0935\u093f\u0935\u0930\u0923"],
        }

        # Replace placeholders
        for para in doc.paragraphs:
            for key, value in replacements.items():
                if key in para.text:
                    para.text = para.text.replace(key, value)

        # Save DOCX
        emp_name = selected_row["\u0915\u0930\u094d\u092e\u091a\u093e\u0930\u0940 \u0915\u093e \u0928\u093e\u092e (\u0939\u093f\u0902\u0926\u0940)"]
        filename_docx = f"SF-11 - {emp_name} - {letter_date}.docx"
        filename_pdf = filename_docx.replace(".docx", ".pdf")
        docx_path = os.path.join(output_folder, filename_docx)
        pdf_path = os.path.join(output_folder, filename_pdf)

        doc.save(docx_path)

        # Try PDF conversion
        if DOCX2PDF_AVAILABLE:
            try:
                convert(docx_path, pdf_path)
                with open(pdf_path, "rb") as f:
                    st.success("✅ Letter generated successfully!")
                    st.download_button("⬇️ Download PDF", f, file_name=filename_pdf)
            except:
                with open(docx_path, "rb") as f:
                    st.warning("⚠️ PDF conversion failed. Download Word file instead.")
                    st.download_button("⬇️ Download Word", f, file_name=filename_docx)
        else:
            with open(docx_path, "rb") as f:
                st.info("ℹ️ PDF not available. Download Word file.")
                st.download_button("⬇️ Download Word", f, file_name=filename_docx)

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
