import streamlit as st
import pandas as pd
from datetime import date, timedelta
from docx import Document
import os
import base64

# === Configurations ===
template_path = "assets/Absent Duty letter temp.docx"
output_folder = "generated"
os.makedirs(output_folder, exist_ok=True)

# === Select Letter Type ===
letter_types = [
    "Duty Letter (For Absent)",
    "SF-11 For Other Reason",
    "Sick Memo",
    "General Letter",
    "Exam NOC",
    "SF-11 Punishment Order"
]
selected_letter_type = st.selectbox("📌 Select Letter Type:", letter_types)

# === Load Employee Data ===
employee_master = pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name=None)
sheet_names = list(employee_master.keys())
selected_sheet = st.selectbox("📋 Select Sheet", sheet_names)
df_emp = employee_master[selected_sheet]
df_emp["Display"] = df_emp.apply(lambda row: f"{row[1]} - {row[2]} - {row[4]} - {row[5]}", axis=1)
selected_display = st.selectbox("👤 Select Employee", df_emp["Display"].dropna().tolist())
selected_row = df_emp[df_emp["Display"] == selected_display].iloc[0]

# === Duty Letter Specific Fields ===
if selected_letter_type == "Duty Letter (For Absent)":
    st.subheader("📄 Generate Duty Letter")

    duty_mode = st.selectbox("🛠 Duty Mode", ["SF-11 & Duty Letter For Absent", "Duty Letter For Absent"])
    from_date = st.date_input("📅 From Date")
    to_date = st.date_input("📅 To Date", value=date.today())
    join_date = st.date_input("📆 Join Date", value=to_date + timedelta(days=1))
    letter_date = st.date_input("📄 Letter Date", value=date.today())

    # Get Fields
    pf_number = selected_row[1]
    hindi_name = selected_row[13]
    designation = selected_row[18]
    short_name = selected_row[14]
    unit = str(selected_row[4])
    working_station = selected_row[8]

    # === Create LetterNo and Memo ===
    unit_code = unit[:2] if len(unit) >= 2 else unit
    letter_no = f"{short_name}/{unit_code}/{working_station}"
    days_absent = (to_date - from_date).days + 1
    memo = f"""आप बिना किसी पूर्व सूचना के दिनांक {from_date.strftime('%d-%m-%Y')} से {to_date.strftime('%d-%m-%Y')} तक कुल {days_absent} दिवस कार्य से अनुपस्थित थे, जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है। अतः आप कामों व भूलो के फेहरिस्त धारा 1, 2 एवं 3 के उल्लंघन के दोषी पाए जाते है।"""

    # === Placeholder Replacement Context ===
    context = {
        "LetterDate": letter_date.strftime("%d-%m-%Y"),
        "EmployeeName": hindi_name,
        "Designation": designation,
        "FromDate": from_date.strftime("%d-%m-%Y"),
        "ToDate": to_date.strftime("%d-%m-%Y"),
        "JoinDate": join_date.strftime("%d-%m-%Y"),
        "PFNumber": pf_number,
        "LetterNo": letter_no,
        "Memo": memo,
"UnitNumber": unit  # <-- नया key
    }

    def generate_word(template_path, context, filename):
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
        save_path = os.path.join(output_folder, filename)
        doc.save(save_path)
        return save_path

    def download_word(path):
        with open(path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        file_name = os.path.basename(path)
        href = f'<a href="data:application/octet-stream;base64,{b64}" download="{file_name}">📥 Download Word File</a>'
        st.markdown(href, unsafe_allow_html=True)

    if st.button("📄 Generate Duty Letter"):
    if duty_mode == "SF-11 & Duty Letter For Absent":
        # 1. SF-11 Generate
        sf11_template = template_files["SF-11 For Other Reason"]
        sf11_filename = f"SF-11 - {hindi_name}.docx"
        sf11_path = generate_doc(sf11_template, context)
        st.success("✅ SF-11 Letter generated successfully!")
        download_file(sf11_path)

        # 2. Duty Letter Generate
        duty_template = template_files["Duty Letter (For Absent)"]
        duty_filename = f"Duty Letter - {hindi_name}.docx"
        duty_path = generate_doc(duty_template, context)
        st.success("✅ Duty Letter generated successfully!")
        download_file(duty_path)

    elif duty_mode == "Duty Letter For Absent":
        duty_template = template_files["Duty Letter (For Absent)"]
        duty_filename = f"Duty Letter - {hindi_name}.docx"
        duty_path = generate_doc(duty_template, context)
        st.success("✅ Duty Letter generated successfully!")
        download_file(duty_path)