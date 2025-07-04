import streamlit as st
from docx import Document
import pandas as pd
from datetime import datetime
import os

# === File paths ===
EMPLOYEE_DATA_PATH = "assets/EMPLOYEE MASTER DATA.xlsx"
DUTY_LETTER_TEMPLATE_PATH = "assets/Absent Duty letter temp.docx"

# === Utility Functions ===
@st.cache_data
def load_employee_master():
    return pd.read_excel(EMPLOYEE_DATA_PATH, sheet_name="Apr.25")

def generate_duty_letter(name, designation, letter_date):
    doc = Document(DUTY_LETTER_TEMPLATE_PATH)
    for para in doc.paragraphs:
        if "[EmployeeName]" in para.text:
            para.text = para.text.replace("[EmployeeName]", name)
        if "[Designation]" in para.text:
            para.text = para.text.replace("[Designation]", designation)
        if "[LetterDate]" in para.text:
            para.text = para.text.replace("[LetterDate]", letter_date.strftime("%d-%m-%Y"))

    filename = f"Duty_Letter_{name.replace(' ', '_')}_{letter_date.strftime('%d%m%Y')}.docx"
    output_path = os.path.join("output", filename)
    os.makedirs("output", exist_ok=True)
    doc.save(output_path)
    return output_path

# === Streamlit UI ===
st.title("📄 Duty Letter Generator")

# Dropdown for letter type
letter_type = st.selectbox("📋 Select Letter Type:", [
    "Duty Letter (For Absent)",
    "SF-11 For Other Reason",
    "Sick Memo",
    "General Letter",
    "Exam NOC",
    "SF-11 Punishment Order"
])

# If Duty Letter
if letter_type == "Duty Letter (For Absent)":
    st.header("📌 Generate Duty Letter for Absent Employee")

    df = load_employee_master()
    df['DisplayName'] = df.apply(lambda row: f"{row[1]} - {row[4]} ({row[5]})", axis=1)  # PF No - Name (Unit)
    selected = st.selectbox("Select Employee:", df['DisplayName'].tolist())

    if selected:
        selected_row = df[df['DisplayName'] == selected].iloc[0]
        emp_name = selected_row[13]       # Column 14 – Hindi Name
        designation = selected_row[17]    # Column 18 – Hindi Designation

        letter_date = st.date_input("📅 Select Letter Date:", value=datetime.today())

        if st.button("📄 Generate Duty Letter"):
            word_path = generate_duty_letter(emp_name, designation, letter_date)
            st.success(f"✅ Duty Letter generated: {os.path.basename(word_path)}")
            with open(word_path, "rb") as f:
                st.download_button("📥 Download Letter", f, file_name=os.path.basename(word_path))