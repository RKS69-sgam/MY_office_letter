
import streamlit as st
import pandas as pd

# Load the employee master Excel file with all sheets
employee_master_file = "assets/EMPLOYEE MASTER DATA.xlsx"
employee_master = pd.read_excel(employee_master_file, sheet_name=None)

# Show sheet selection dropdown
sheet_names = list(employee_master.keys())
selected_sheet = st.selectbox("2ï¸âƒ£ Select Sheet (Employee Master):", sheet_names)

# Load the selected sheet's data
df_emp = employee_master[selected_sheet]

# Generate Display column
df_emp["Display"] = df_emp.apply(
    lambda row: f"{row[1]} - {row[2]} - {row[4]} - {row[5]}", axis=1
)

# Show employee selection dropdown
emp_display_list = df_emp["Display"].dropna().tolist()
selected_emp_display = st.selectbox("3ï¸âƒ£ Select Employee:", emp_display_list)

# Extract details for selected employee
if selected_emp_display:
    selected_row = df_emp[df_emp["Display"] == selected_emp_display].iloc[0]
    pf_number = selected_row[1]
    hrms_id = selected_row[2]
    unit_raw = selected_row[4]
    working_station = selected_row[8]
    english_name = selected_row[5]
    hindi_name = selected_row[13]
    designation = selected_row[18]
    short_name = selected_row[14] if len(selected_row) > 14 else ""

# === Duty Letter (For Absent) Section ===
if selected_letter_type == "Duty Letter (For Absent)":
    st.subheader("ðŸ“„ Duty Letter (For Absent)")

    duty_mode = st.selectbox("ðŸ“Œ Duty Letter Type:", [
        "SF-11 & Duty Letter For Absent",
        "Duty Letter For Absent"
    ])

    from_date = st.date_input("ðŸ—“ From Date")
    to_date = st.date_input("ðŸ—“ To Date", date.today())
    join_date = to_date + timedelta(days=1)
    duty_join_date = st.date_input("ðŸ“† Join Date", join_date)

    # Prepare context
    context = {
        "LetterDate": letter_date.strftime("%d-%m-%Y"),
        "EmployeeName": hindi_name,
        "Designation": designation,
        "PFNumber": pf_number,
        "FromDate": from_date.strftime("%d-%m-%Y"),
        "ToDate": to_date.strftime("%d-%m-%Y"),
        "JoinDate": duty_join_date.strftime("%d-%m-%Y"),
        "LetterNo": f"{short_name} / {str(unit_raw)[:2]} / {working_station}"
    }

    if st.button("ðŸ“„ Generate Duty Letter"):
        doc = Document(template_files["Duty Letter (For Absent)"])
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

        file_name = f"Duty_Letter_{english_name}_{letter_date.strftime('%d-%m-%Y')}"
        output_path = os.path.join("/tmp", f"{file_name}.docx")
        doc.save(output_path)

        st.success("âœ… Duty Letter Generated Successfully!")
        with open(output_path, "rb") as f:
            st.download_button("â¬‡ï¸ Download Word File", f, file_name + ".docx")

        pdf_path = convert_to_pdf(output_path)
        if pdf_path and os.path.exists(pdf_path):
            with open(pdf_path, "rb") as f:
                st.download_button("â¬‡ï¸ Download PDF File", f, os.path.basename(pdf_path))
        else:
            st.warning("âš ï¸ PDF conversion failed or not supported.")