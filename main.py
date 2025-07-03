import streamlit as st
import pandas as pd
import os
from docx import Document
from datetime import datetime

# === Load SF-11 Register (Live Read with No Cache) ===
@st.cache_data(ttl=0)
def load_sf11_register():
    return pd.read_excel("SF-11 Register.xlsx", sheet_name="SSE-SGAM")

# === Load Master Data (if needed for other letter types) ===
@st.cache_data(ttl=0)
def load_master_data():
    return pd.read_excel("EMPLOYEE MASTER DATA.xlsx", sheet_name="Apr.25")

# === Mapping for Templates ===
template_files = {
    "SF-11 Punishment Order": "assets/SF-11 Punishment Order Template.docx",
    "SF-11 For Other Reason": "assets/SF-11 Other Reason Template.docx",
    "Duty Letter (For Absent)": "assets/Duty Letter Template.docx",
    "Sick Memo": "assets/Sick Memo Template.docx",
    "Exam NOC": "assets/Exam NOC Template.docx",
    "General Letter": "assets/General Letter Template.docx"
}

# === Letter Type Dropdown ===
st.title("📂 SF-11 & Other Letters Generator")
letter_type = st.selectbox("1️⃣ Select Letter Type:", list(template_files.keys()))

# === Load relevant data based on Letter Type ===
if letter_type == "SF-11 Punishment Order":
    df = load_sf11_register()
    df["Display"] = df["PF No."].astype(str) + " - " + df["Employee Name"] + " - " + df["Letter Date"].astype(str)
    selected_display = st.selectbox("👤 Select Employee (SF-11 Register):", df["Display"])
    selected_row = df[df["Display"] == selected_display].iloc[0]

    letter_date = st.date_input("📅 Letter Date", datetime.today())

    if st.button("📄 Generate SF-11 Punishment Order"):
        doc = Document(template_files[letter_type])

        # Replace placeholders
        for para in doc.paragraphs:
            if "[EmployeeName]" in para.text:
                para.text = para.text.replace("[EmployeeName]", selected_row["Employee Name"])
            if "[Designation]" in para.text:
                para.text = para.text.replace("[Designation]", selected_row["Designation"])
            if "[LetterDate]" in para.text:
                para.text = para.text.replace("[LetterDate]", letter_date.strftime("%d-%m-%Y"))
            if "[LetterNo.]" in para.text:
                para.text = para.text.replace("[LetterNo.]", selected_row["पत्र क्र."])
            if "[MEMO]" in para.text:
                para.text = para.text.replace("[MEMO]", selected_row["Memo"])

        output_path = f"generated/SF-11 - {selected_row['Employee Name']}.docx"
        os.makedirs("generated", exist_ok=True)
        doc.save(output_path)
        st.success("✅ Letter generated successfully!")

        with open(output_path, "rb") as f:
            st.download_button("⬇️ Download Word File", f, file_name=os.path.basename(output_path))
