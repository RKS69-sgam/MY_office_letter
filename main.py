import streamlit as st
import pandas as pd
from docx import Document
from datetime import datetime
import os

# === File paths ===
MASTER_DATA_PATH = "EMPLOYEE MASTER DATA.xlsx"
TEMPLATE_PATH = "Duty Letter Template.docx"
OUTPUT_DIR = "Generated_Letters"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# === Load employee data ===
@st.cache_data
def load_employee_master():
    df = pd.read_excel(MASTER_DATA_PATH, sheet_name="Apr.25")
    df = df.dropna(subset=["PF Number", "Name (Hindi)", "Designation (Hindi)"])
    return df

# === Main app ===
st.set_page_config(page_title="Letter Generator", layout="centered")
st.title("📄 Railway Letter Generator")

# Step 1: Select letter type
letter_type = st.selectbox(
    "📑 Select Letter Type:",
    [
        "Duty Letter (For Absent)",
        "SF-11 For Other Reason",
        "Sick Memo",
        "General Letter",
        "Exam NOC",
        "SF-11 Punishment Order"
    ]
)

# === Duty Letter (For Absent) ===
if letter_type == "Duty Letter (For Absent)":
    st.header("📌 Generate Duty Letter for Absent Employee")

    df = load_employee_master()

    emp_options = df.apply(lambda row: f"{row['PF Number']} - {row['Name (Hindi)']} - {row['Designation (Hindi)']}", axis=1)
    selected = st.selectbox("👤 Select Employee", emp_options)

    letter_date = st.date_input("📅 Letter Date", value=datetime.today())

    if st.button("📄 Generate Duty Letter"):
        idx = emp_options[emp_options == selected].index[0]
        emp = df.loc[idx]

        replacements = {
            "[Name]": emp["Name (Hindi)"],
            "[Designation]": emp["Designation (Hindi)"],
            "[LetterDate]": letter_date.strftime("%d-%m-%Y"),
        }

        try:
            doc = Document(TEMPLATE_PATH)
            for para in doc.paragraphs:
                for key, val in replacements.items():
                    if key in para.text:
                        para.text = para.text.replace(key, val)

            filename = f"Duty Letter - {emp['Name (Hindi)']} - {letter_date.strftime('%Y%m%d')}.docx"
            output_path = os.path.join(OUTPUT_DIR, filename)
            doc.save(output_path)

            st.success("✅ Duty Letter generated successfully!")
            with open(output_path, "rb") as f:
                st.download_button("⬇️ Download Letter", f, file_name=filename)

        except Exception as e:
            st.error(f"❌ Error generating letter: {e}")