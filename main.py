import streamlit as st
import pandas as pd
from datetime import date, timedelta
from docx import Document
from tempfile import NamedTemporaryFile
import base64
import os
import shutil

# === Template file mapping (updated) ===
template_files = {
    "SF-11 Punishment Order": "assets/SF-11 Punishment order temp.docx",
    "SF-11 For Other Reason": "assets/SF-11 temp.docx",
    "Duty Letter (For Absent)": "assets/Absent Duty letter temp.docx",
    "Sick Memo": "assets/SICK MEMO temp..docx",
    "Exam NOC": "assets/Exam NOC Letter temp.docx",
    "General Letter": "assets/General Letter temp.docx"
}

# === Placeholder functions for next sections ===
def load_master_data():
    # Load EMPLOYEE MASTER DATA and SF-11 Register
    pass

def load_ui():
    # Render UI and collect inputs according to letter type
    pass

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

def download_file(file_path):
    with open(file_path, "rb") as f:
        data = f.read()
        b64 = base64.b64encode(data).decode()
        href = f'<a href="data:application/octet-stream;base64,{b64}" download="{os.path.basename(file_path)}">Download File</a>'
        st.markdown(href, unsafe_allow_html=True)

# === Main App Logic ===
# This will be built out step-by-step according to full logic
st.title("Letter Generator with SF-11 & General Letter Support")
st.markdown("---")
st.success("Template files mapped correctly and ready to use.")
st.success("Template files mapped correctly and ready to use.")

# === Select Letter Type ===
letter_types = [
    "SF-11 Punishment Order",
    "SF-11 For Other Reason",
    "Duty Letter (For Absent)",
    "Sick Memo",
    "Exam NOC",
    "General Letter"
]
selected_letter_type = st.selectbox("1️⃣ Select Letter Type:", letter_types)

# === Load Required Data ===
employee_master = pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name=None)
sf11_register = pd.read_excel("assets/SF-11 Register.xlsx", sheet_name="SSE-SGAM")
exam_noc_data = pd.read_excel("assets/ExamNOC_Report.xlsx")

# === Conditional Sheet Source ===
if selected_letter_type == "SF-11 Punishment Order":
    # Show dropdown from SF-11 Register
    sf11_register["Display"] = sf11_register.apply(
        lambda row: f"{row['पी.एफ. क्रमांक']} - {row['कर्मचारी का नाम']} - {row['दिनांक']} - {row['पत्र क्र.']}",
        axis=1
    )
    sf11_employee_list = sf11_register["Display"].tolist()
    selected_sf11_display = st.selectbox("2️⃣ Select Employee (SF-11 Register):", sf11_employee_list)
    selected_sf11_row = sf11_register[sf11_register["Display"] == selected_sf11_display].iloc[0]

else:
    # Show dropdown from Employee Master Data
    sheet_names = list(employee_master.keys())
    selected_sheet = st.selectbox("2️⃣ Select Sheet (Employee Master):", sheet_names)
    df_emp = employee_master[selected_sheet]
    df_emp["Display"] = df_emp.apply(
        lambda row: f"{row[1]} - {row[2]} - {row[4]} - {row[5]}", axis=1
    )
    emp_display_list = df_emp["Display"].dropna().tolist()
    selected_emp_display = st.selectbox("3️⃣ Select Employee:", emp_display_list)
    selected_row = df_emp[df_emp["Display"] == selected_emp_display].iloc[0]
pf_number = selected_row[1]
    hrms_id = selected_row[2]
    unit_raw = selected_row[4]
    working_station = selected_row[8]
    english_name = selected_row[5]
    hindi_name = selected_row[13]
    designation = selected_row[18]
    short_name = selected_row[14] if len(selected_row) > 14 else ""

    # Letter date
    letter_date = st.date_input("📅 Letter Date", date.today())

    # === Duty Letter Specific ===
    if selected_letter_type == "Duty Letter (For Absent)":
        duty_mode = st.selectbox("📌 Duty Letter Type:", ["SF-11 & Duty Letter For Absent", "Duty Letter For Absent"])
        from_date = st.date_input("🗓 From Date")
        to_date = st.date_input("🗓 To Date", date.today())
        join_date = to_date + timedelta(days=1)
        duty_join_date = st.date_input("📆 Join Date", join_date)

    # === SF-11 For Other Reason ===
    if selected_letter_type == "SF-11 For Other Reason":
        memo_text = st.text_area("📝 Memorandum Text")

    # === Exam NOC ===
    if selected_letter_type == "Exam NOC":
        exam_year = st.selectbox("📆 NOC Year", [date.today().year])
        exam_name = st.text_input("🧪 Exam Name")
        emp_past_nocs = exam_noc_data[(exam_noc_data["PF Number"] == pf_number) & (exam_noc_data["NOC Year"] == exam_year)]
        taken_count = len(emp_past_nocs)
        if taken_count >= 4:
            st.warning(f"⚠️ Already taken {taken_count} NOCs in {exam_year}. Cannot apply more.")
        else:
            next_noc = taken_count + 1
            noc_number = st.selectbox("🔢 NOC Attempt No.", [next_noc, *range(next_noc + 1, 5)])

    # === General Letter ===
    if selected_letter_type == "General Letter":
        officer_list = [
            "सहायक मण्डल अभियंता\nब्यौहारी",
            "मंडल रेल प्रबंधक (कार्मिक)\nप. म. रे. जबलपुर",
            "मण्डल अभियंता (पूर्व)\nप. म. रे. जबलपुर मण्डल",
            "मेट यू क्रमांक 30",
            "Other"
        ]
        officer_to = st.selectbox("📌 Letter To", officer_list)
        if officer_to == "Other":
            officer_to = st.text_area("✍ Enter Other Officer")

        subject_text = st.text_input("📄 Subject")
        reference_text = st.text_area("📎 Reference (optional)")
        detail_memo = st.text_area("📝 Detailed Memo (Justified)", height=150)
        copy_to_options = [
            "",
            "सहायक मण्डल अभियंता ब्योहारी को सूचनार्थ सादर संप्रेषित ।",
            "मंडल रेल प्रबंधक (कार्मिक) प. म. रे. जबलपुर को सूचनार्थ सादर संप्रेषित ।",
            "मण्डल अभियंता (पूर्व) प. म. रे. जबलपुर मण्डल को सूचनार्थ सादर संप्रेषित ।"
        ]
        copy_to = st.multiselect("📋 Copy To (Optional)", copy_to_options)
# Prepare General Letter Context
    context["LetterTo"] = officer_to if officer_to != "Other" else custom_officer_to
    context["Subject"] = subject_text
    context["Reference"] = reference_text if reference_text.strip() else ""
    context["DetailMemo"] = detail_memo
    context["CopyTo"] = "\n".join(copy_to) if copy_to else ""

    # Hide reference and copy-to block if empty
    if not context["Reference"] and not context["CopyTo"]:
        context["HideRef"] = "yes"
    else:
        context["HideRef"] = "no"

    if st.button("📄 Generate General Letter"):
        doc = Document(template_files["General Letter"])
        replace_placeholders(doc, context)

        # Handle Reference and CopyTo block visibility
        if context["HideRef"] == "yes":
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if "[Reference]" in cell.text:
                            cell.text = ""
                        if "[CopyTo]" in cell.text:
                            cell.text = ""

        file_name = f"General_Letter_{english_name}_{letter_date.strftime('%d-%m-%Y')}"
        output_docx = os.path.join("/tmp", f"{file_name}.docx")
        doc.save(output_docx)

        st.success("✅ General Letter Generated Successfully!")
        download_button(output_docx, f"⬇️ Download {file_name}.docx")

        pdf_path = convert_to_pdf(output_docx)
        if pdf_path and os.path.exists(pdf_path):
            st.success("📄 PDF also generated!")
            download_button(pdf_path, f"⬇️ Download {os.path.basename(pdf_path)}")
        else:
            st.warning("⚠️ PDF conversion failed or not supported.")
elif selected_type == "Exam NOC":
    st.subheader("📝 Exam NOC Letter")

    # Year and attempt selection
    current_year = date.today().year
    noc_year = st.selectbox("📅 Select NOC Year", [current_year])
    exam_name = st.text_input("📝 Exam Name")

    # Get previous NOC attempts from Excel
    exam_df = pd.read_excel("assets/ExamNOC_Report.xlsx")
    emp_noc_count = exam_df[
        (exam_df["PF Number"] == selected_row[col_pf]) & 
        (exam_df["NOC Year"] == noc_year)
    ].shape[0]

    if emp_noc_count >= 4:
        st.error(f"❌ Already taken 4 NOCs in {noc_year}.")
    else:
        next_attempt = emp_noc_count + 1
        application_no = st.selectbox("🔢 NOC Attempt No", [next_attempt])

        # Add employee entry to table list
        if st.button("➕ Add to NOC List"):
            new_entry = {
                "PF Number": selected_row[col_pf],
                "Employee Name": hindi_name,
                "Designation": selected_row[col_designation],
                "NOC Year": noc_year,
                "Application No.": application_no,
                "Exam Name": exam_name
            }
            exam_df = pd.concat([exam_df, pd.DataFrame([new_entry])], ignore_index=True)
            exam_df.to_excel("assets/ExamNOC_Report.xlsx", index=False)
            st.success("✅ Employee NOC entry added.")

        # Display the list in table format
        st.write("📋 NOC Application List")
        st.dataframe(
            exam_df[
                (exam_df["PF Number"] == selected_row[col_pf]) & 
                (exam_df["NOC Year"] == noc_year)
            ][["PF Number", "Employee Name", "Designation", "NOC Year", "Application No.", "Exam Name"]]
        )

        # Table for template
        from docx.shared import Inches
        table_rows = exam_df[
            (exam_df["PF Number"] == selected_row[col_pf]) & 
            (exam_df["NOC Year"] == noc_year)
        ][["PF Number", "Employee Name", "Designation", "NOC Year", "Application No.", "Exam Name"]].values.tolist()

        # Word file generation
        if st.button("📄 Generate Exam NOC Letter"):
            doc = Document(template_files["Exam NOC"])

            # Fill placeholders
            replace_placeholders(doc, {
                "LetterDate": letter_date.strftime("%d-%m-%Y"),
                "EmployeeName": hindi_name,
                "Designation": selected_row[col_designation],
                "PFNumber": "",  # Placeholder for table
                "ExamName": exam_name,
                "NOCCount": next_attempt
            })

            # Create table in placeholder cell
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if "[PFNumber]" in cell.text:
                            cell.text = ""
                            tbl = cell.add_table(rows=1, cols=6)
                            tbl.style = 'Table Grid'
                            hdr_cells = tbl.rows[0].cells
                            hdr_cells[0].text = "PF Number"
                            hdr_cells[1].text = "Employee Name"
                            hdr_cells[2].text = "Designation"
                            hdr_cells[3].text = "NOC Year"
                            hdr_cells[4].text = "Application No."
                            hdr_cells[5].text = "Exam Name"

                            for data_row in table_rows:
                                row_cells = tbl.add_row().cells
                                for i in range(6):
                                    row_cells[i].text = str(data_row[i])
                            break

            file_name = f"ExamNOC_{english_name}_{letter_date.strftime('%d-%m-%Y')}"
            output_docx = os.path.join("/tmp", f"{file_name}.docx")
            doc.save(output_docx)

            st.success("✅ Exam NOC Letter Generated!")
            download_button(output_docx, f"⬇️ Download {file_name}.docx")

            pdf_path = convert_to_pdf(output_docx)
            if pdf_path and os.path.exists(pdf_path):
                st.success("📄 PDF also generated!")
                download_button(pdf_path, f"⬇️ Download {os.path.basename(pdf_path)}")
            else:
                st.warning("⚠️ PDF conversion failed or not supported.")
elif selected_type == "SF-11 For Other Reason":
    st.subheader("📄 SF-11 Letter for Other Reason")

    # Select employee from Employee Master Data
    emp_display_list = df["DisplayName"].dropna().tolist()
    selected_emp_display = st.selectbox("👤 Select Employee", emp_display_list)

    if selected_emp_display:
        selected_row = df[df["DisplayName"] == selected_emp_display].iloc[0]
        english_name = selected_row[col_english_name]
        hindi_name = selected_row[col_hindi_name]
        pf_number = selected_row[col_pf]
        designation = selected_row[col_designation]

        # Memorandum textbox
        memo = st.text_area("📝 Memorandum")

        # Auto format final memo
        full_memo = f"{memo}, जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है। अतः आप कामों व भूलो के फेहरिस्त धारा 1, 2 एवं 3 के उल्लंघन के दोषी पाए जाते है।"

        # Generate button
        if st.button("📄 Generate SF-11 Letter"):
            context = {
                "LetterDate": letter_date.strftime("%d-%m-%Y"),
                "EmployeeName": hindi_name,
                "Designation": designation,
                "Memo": full_memo,
                "PFNumber": pf_number,
            }

            filename = f"SF11_OtherReason_{english_name}_{letter_date.strftime('%d-%m-%Y')}"
            output_path = generate_docx(template_files["SF-11 For Other Reason"], context, filename)
            st.success("✅ SF-11 Letter Generated!")
            download_button(output_path, f"⬇️ Download {os.path.basename(output_path)}")

            # Save to SF-11 Register
            try:
                sf11_wb = pd.read_excel("assets/SF-11 Register.xlsx", sheet_name="SSE-SGAM")
                last_index = sf11_wb.shape[0]
                new_entry = {
                    "स.क्र.": last_index + 1,
                    "पी.एफ. क्रमांक": pf_number,
                    "कर्मचारी का नाम": hindi_name,
                    "पदनाम": designation,
                    "पत्र क्र.": f"{selected_row[15]} / {str(selected_row[col_unit])[:2]} / {selected_row[9]}",
                    "दिनांक": letter_date.strftime("%d-%m-%Y"),
                    "आरोप का विवरण": full_memo
                }
                updated_sf11 = pd.concat([sf11_wb, pd.DataFrame([new_entry])], ignore_index=True)
                updated_sf11.to_excel("assets/SF-11 Register.xlsx", sheet_name="SSE-SGAM", index=False)
                st.success("🗃️ Entry saved to SF-11 Register.")
            except Exception as e:
                st.error(f"❌ Error saving to register: {e}")

            pdf_path = convert_to_pdf(output_path)
            if pdf_path and os.path.exists(pdf_path):
                st.success("📄 PDF also generated!")
                download_button(pdf_path, f"⬇️ Download {os.path.basename(pdf_path)}")
            else:
                st.warning("⚠️ PDF conversion failed or not supported.")
