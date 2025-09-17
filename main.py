import streamlit as st
import pandas as pd
import os
import base64
from docx import Document
from datetime import date, timedelta
import datetime
from docx.shared import Inches
from io import BytesIO

# Create output folder (This is for local development; it will not work in the cloud)
os.makedirs("generated_letters", exist_ok=True)

# File paths
template_files = {
    "Duty Letter (For Absent)": "assets/Absent Duty letter temp.docx",
    "SF-11 For Other Reason": "assets/SF-11 temp.docx",
    "Sick Memo": "assets/SICK MEMO temp..docx",
    "General Letter": "assets/General Letter temp.docx",
    "Exam NOC": "assets/Exam NOC Letter temp.docx",
    "SF-11 Punishment Order": "assets/SF-11 Punishment order temp.docx",
    "Quarter Allotment Letter": "assets/Quarter Allotment temp.docx",
    "Update Employee Database": None,
    "Engine Pass Letter": "assests/Engine Pass letter temp.docx",
    "Card Pass Letter": "assests/Card Pass letter temp.docx",
    "DAR NOC Letter": "assets/DAR NOC temp.docx"
}
quarter_file = "assets/QUARTER REGISTER.xlsx"
if os.path.exists(quarter_file):
    quarter_df = pd.read_excel(quarter_file, sheet_name="Sheet1")
else:
    quarter_df = pd.DataFrame() # Fallback for cloud deployment
employee_master = pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name=None)
sf11_register_path = "assets/SF-11 Register.xlsx"
if os.path.exists(sf11_register_path):
    sf11_register = pd.read_excel(sf11_register_path, sheet_name="SSE-SGAM")
else:
    sf11_register = pd.DataFrame()
noc_register_path = "assets/Exam NOC_Report.xlsx"
df_noc = pd.read_excel(noc_register_path) if os.path.exists(noc_register_path) else pd.DataFrame(columns=)

# Placeholder replace in paragraph
def replace_placeholder_in_para(paragraph, context):
    full_text = ''.join(run.text for run in paragraph.runs)
    new_text = full_text
    for key, val in context.items():
        new_text = new_text.replace(f"[{key}]", str(val))
    if new_text!= full_text:
        for run in paragraph.runs:
            run.text = ''
        if paragraph.runs:
            paragraph.runs.text = new_text
        else:
            paragraph.add_run(new_text)

# === Generate Word Function (Refactored for In-Memory operation) ===
def generate_word(template_path, context):
    doc = Document(template_path)
    # Replace in paragraphs
    for p in doc.paragraphs:
        replace_placeholder_in_para(p, context)
    # Replace in table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_placeholder_in_para(p, context)

    # ✅ Exam NOC Table Insertion
    if context.get("LetterType") == "Exam NOC":
        for i, paragraph in enumerate(doc.paragraphs):
            if "[PFNumber]" in paragraph.text:
                p = paragraph._element
                p.getparent().remove(p)
                p._p = p._element = None
                table = doc.add_table(rows=1, cols=6)
                table.style = "Table Grid"
                table.autofit = True
                hdr = table.rows.cells
                hdr.text = "PF Number"
                hdr.[1]text = "Employee Name"
                hdr.[2]text = "Designation"
                hdr.[3]text = "NOC Year"
                hdr.[4]text = "Application No."
                hdr.[5]text = "Exam Name"
                row = table.add_row().cells
                row.text = str(context["PFNumberVal"])
                row.[1]text = context["EmployeeName"]
                row.[2]text = context
                row.[3]text = str(context)
                row.[4]text = str(context["AppNo"])
                row.[5]text = context["ExamName"]
                break
    
    # Use BytesIO to save the document to an in-memory buffer
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# === Download Function (Not used directly, but part of old logic) ===
def download_word(path):
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
        name = os.path.basename(path)
        href = f'<a href="data:application/octet-stream;base64,{b64}" download="{name}">Download Word File</a>'
        st.markdown(href, unsafe_allow_html=True)

# Placeholder for handle_engine_card_pass function
def handle_engine_card_pass(letter_type):
    st.warning(f"The `handle_engine_card_pass` function for '{letter_type}' is not included in this script. Please ensure it's imported or defined elsewhere.")
    if letter_type in ["Engine Pass Letter", "Card Pass Letter"]:
        return "N/A"
    return ""

# === UI ===
st.title("OFFICE OF THE SSE/PW/SGAM")

# Password protection
password = st.text_input("Enter Password", type="password")
if password == "sgam@4321":
    st.success("Access Granted!")

    letter_type = st.selectbox("Select Letter Type", list(template_files.keys()))

    # === Select Employee Logic ===
    dor_str = ""
    if letter_type in ["Engine Pass Letter", "Card Pass Letter"]:
        dor_str = handle_engine_card_pass(letter_type)
        df = employee_master["Apr.25"]
        df = df.apply(lambda r: f"{r[1]} - {r[2]} - {r[4]} - {r[5]}", axis=1)
        selected = st.selectbox("Select Employee", df.dropna())
        row = df == selected].iloc
        pf = row[1]
        hname = row
        desg = row
        unit_full = str(row[4])
        unit = unit_full[:2]
        short = row
        letter_no = f"{short}/{unit}/{unit}"
        letter_date = st.date_input("Letter Date", value=date.today())
    elif letter_type == "SF-11 Punishment Order":
        df = sf11_register
        df = df.apply(lambda r: f"{r['पी.एफ. क्रमांक']} - {r['कर्मचारी का नाम']} - {r['पत्र क्र.']} - {r['दिनांक']}", axis=1)
        selected = st.selectbox("Select Employee", df.dropna())
        row = df == selected].iloc
        patra_kr = row["पत्र क्र."]
        dandadesh_krmank = f"{patra_kr}/D-1"
        pf = row["पी.एफ. क्रमांक"]
        hname = row["कर्मचारी का नाम"]
        desg = row.get("पदनाम", "")
        unit_full = patra_kr.split("/", 1)[1]
        unit = unit_full[-7:]
        short = patra_kr.split("/")
        letter_no = dandadesh_krmank
        sf11date = row["दिनांक"]
        letter_date = st.date_input("Letter Date", value=date.today())
    elif letter_type == "General Letter":
        df = pd.DataFrame()
        pf = hname = desg = unit = unit_full = short = letter_no = ""
        letter_date = st.date_input("Letter Date", value=date.today())
    elif letter_type == "Update Employee Database":
        pass
    else: # For Duty Letter, SF-11 For Other Reason, Sick Memo, Exam NOC, Quarter Allotment, and DAR NOC Letter
        df = employee_master["Apr.25"]
        df = df.apply(lambda r: f"{r[1]} - {r[2]} - {r[4]} - {r[5]}", axis=1)
        selected = st.selectbox("Select Employee", df.dropna())
        row = df == selected].iloc
        pf = row[1]
        hname = row
        desg = row
        unit_full = str(row[4])
        unit = unit_full[:2]
        short = row
        letter_no = f"{short}/{unit}/{unit}"
        letter_date = st.date_input("Letter Date", value=date.today())

    # === Common context for template replacement ===
    context = {
        "LetterDate": letter_date.strftime("%d-%m-%Y") if 'letter_date' in locals() else "",
        "Date": letter_date.strftime("%d-%m-%Y") if 'letter_date' in locals() else "",
        "EmployeeName": hname if 'hname' in locals() else "",
        "Designation": desg if 'desg' in locals() else "",
        "PFNumber": pf if 'pf' in locals() else "",
        "ShortName": short if 'short' in locals() else "",
        "Unit": unit if 'unit' in locals() else "",
        "UnitNumber": unit if 'unit' in locals() else "",
        "LetterNo": letter_no if 'letter_no' in locals() else "",
        "DutyDate": "",
        "FromDate": "",
        "ToDate": "",
        "JoinDate": "",
        "Memo": "",
        "OfficerUnit": "",
        "Subject": "",
        "Reference": "",
        "CopyTo": "",
        "DOR": dor_str
    }

    if letter_type == "Duty Letter (For Absent)":
        mode = st.selectbox("Mode",)
        fd = st.date_input("From Date")
        td = st.date_input("To Date", value=date.today())
        jd = st.date_input("Join Date", value=td + timedelta(days=1))
        context.update({
            "FromDate": fd.strftime("%d-%m-%Y"),
            "ToDate": td.strftime("%d-%m-%Y"),
            "JoinDate": jd.strftime("%d-%m-%Y"),
            "DutyDate": jd.strftime("%d-%m-%Y"),
            "Memo": f"आप बिना किसी पूर्व सूचना के दिनांक {fd.strftime('%d-%m-%Y')} से {td.strftime('%d-%m-%Y')} तक कुल {(td-fd).days+1} दिवस कार्य से अनुपस्थित थे, जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है। अतः आप कामों व भूलो के फेहरिस्त धारा 1, 2 एवं 3 के उल्लंघन के दोषी पाए जाते है।"
        })
    elif letter_type == "SF-11 For Other Reason":
        memo_input = st.text_area("Memo")
        context["Memo"] = memo_input + " जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है। अतः आप कामों व भूलो के फेहरिस्त धारा 1, 2 एवं 3 के उल्लंघन के दोषी पाए जाते है।"
    elif letter_type == "General Letter":
        context["FileName"] = st.selectbox("File Name",)
        officer_option = st.selectbox("अधिकारी/कर्मचारी",)
        if officer_option == "अन्‍य":
            officer_option = st.text_input("अन्‍य का नाम/पदनाम/एजेंसी का नाम लिखें")
        context["OfficerName"] = officer_option
        beyohari_officers =
        jbp_officers =
        if officer_option == "कनिष्‍ठ अभियंता (रेल पथ)":
            address_choices = ["निवासरोड", "भरसेड़ी", "गजराबहरा", "गोंदवाली", "अन्‍य"]
        elif officer_option in beyohari_officers:
            address_choices = ["प.म.रे. ब्‍योहारी", "अन्‍य"]
        elif officer_option in jbp_officers:
            address_choices = ["प.म.रे. जबलपुर", "अन्‍य"]
        else:
            address_choices = ["", "प.म.रे. ब्‍योहारी", "प.म.रे. जबलपुर", "सरईग्राम", "देवराग्राम", "बरगवॉं", "निवासरोड", "भरसेड़ी", "गजराबहरा", "गोंदवाली", "अन्‍य"]
        address_option = st.selectbox("पता", address_choices)
        if address_option == "अन्‍य":
            address_option = st.text_input("अन्‍य का पता लिखें")
        context["OfficeAddress"] = address_option

        subject_input = st.text_input("विषय")
        context = f"विषय:- {subject_input}" if subject_input.strip() else ""
        ref_input = st.text_input("संदर्भ")
        context = f"संदर्भ:- {ref_input}" if ref_input.strip() else ""
        context["Memo"] = st.text_area("मुख्‍य विवरण")
        copy_input = st.text_input("प्रतिलिपि")
        context = f"प्रतिलिपि:- " + "\n".join(
            [c.strip() for c in copy_input.split(",") if c.strip()]
        ) if copy_input.strip() else ""

    elif letter_type == "Exam NOC":
        year = date.today().year
        df_match = df_noc[(df_noc["PF Number"] == pf) & (df_noc == year)]
        count = df_match.shape
        if count >= 4:
            st.warning("यह कर्मचारी इस वर्ष पहले ही 4 NOC ले चुका है।")
        else:
            exam_name = st.text_input("Exam Name", key="exam_name")
            term = st.text_input("Term of NOC", key="noc_term")
            context.update({
                "PFNumberVal": pf,
                "EmployeeName": hname,
                "Designation": desg,
                "NOCYear": year,
                "AppNo": count + 1,
                "ExamName": exam_name,
                "Term": term,
                "LetterType": "Exam NOC"
            })
    elif letter_type == "SF-11 Punishment Order":
        st.markdown("#### SF-11 Register से विवरण")
        st.markdown(f"**आरोप का विवरण:** {row.get('आरोप का विवरण', '—')}")
        pawati_date = st.date_input("पावती का दिनांक", value=date.today())
        pratyuttar_date = st.date_input("यदि प्रत्‍युत्तर प्राप्‍त हुआ हो तो दिनांक", value=date.today())
        context["Memo"] = st.selectbox("Punishment Type",)
        context = letter_no
        context["LetterNo."] = patra_kr
        context["Unit"] = unit
        context = sf11date
    elif letter_type == "Quarter Allotment Letter":
        pf = row[1]
        hname = row
        desg = row
        unit_full = str(row[4])
        unit = unit_full[:2]
        quarter_df = quarter_df.apply(lambda r: f"{r} - {r}", axis=1)
        q_selected = st.selectbox("Select Quarter", quarter_df.dropna())
        qrow = quarter_df == q_selected].iloc
        station = qrow
        qno = qrow
        context = {
            "EmployeeName": hname,
            "Designation": desg,
            "Unit": unit,
            "LetterDate": letter_date.strftime("%d-%m-%Y"),
            "QuarterNo.": qno,
            "Station": station
        }
    elif letter_type == "DAR NOC Letter":
        st.info("No additional input required. All data will be pulled from the employee master.")

    elif letter_type == "Update Employee Database":
        st.subheader("Update Employee Database")
        emp_df = employee_master["Apr.25"]
        headers = list(emp_df.columns)
        if "Remark" not in emp_df.columns:
            emp_df = ""
        date_fields =
        action = st.radio("Select Action",)

        if action == "Add New Employee":
            st.subheader("Add New Employee")
            new_data = {}
            for col in headers[:-1]:
                if col in date_fields:
                    new_data[col] = st.date_input(col, key=f"add_{col}")
                else:
                    new_data[col] = st.text_input(col, key=f"add_{col}")
            if st.button("Add Employee"):
                for col in date_fields:
                    if isinstance(new_data[col], date):
                        new_data[col] = new_data[col].strftime("%d-%m-%Y")
                new_data = "Added"
                emp_df = pd.concat()], ignore_index=True)
                employee_master["Apr.25"] = emp_df
                with pd.ExcelWriter("assets/EMPLOYEE MASTER DATA.xlsx", engine="openpyxl") as writer:
                    for sheet, df in employee_master.items():
                        df.to_excel(writer, sheet_name=sheet, index=False)
                st.success(f"Employee added successfully at row {emp_df.shape}.")
        elif action == "Update Existing Employee":
            st.subheader("Update Existing Employee")
            pf_list = emp_df["PF No."].dropna().unique()
            selected_pf = st.selectbox("Select PF Number", pf_list, key="upd_pf")
            if selected_pf:
                row = emp_df[emp_df["PF No."] == selected_pf].iloc
                updated_data = {}
                for col in headers[:-1]:
                    if col in date_fields:
                        date_val = pd.to_datetime(row[col], errors="coerce")
                        updated_data[col] = st.date_input(col, value=date_val if pd.notna(date_val) else date.today(), key=f"upd_{col}")
                    else:
                        updated_data[col] = st.text_input(col, value=row[col], key=f"upd_{col}")
                if st.button("Update Employee"):
                    index = emp_df[emp_df["PF No."] == selected_pf].index
                    for col in headers[:-1]:
                        val = updated_data[col]
                        if col in date_fields and isinstance(val, date):
                            val = val.strftime("%d-%m-%Y")
                        emp_df.at[index, col] = val
                    emp_df.at = "Updated"
                    employee_master["Apr.25"] = emp_df
                    with pd.ExcelWriter("assets/EMPLOYEE MASTER DATA.xlsx", engine="openpyxl") as writer:
                        for sheet, df in employee_master.items():
                            df.to_excel(writer, sheet_name=sheet, index=False)
                    st.success(f"Employee updated at row {index + 1}.")
        elif action == "Mark as Exited (Transfer)":
            st.subheader("Mark Employee as Exited")
            pf_list = emp_df["PF No."].dropna().unique()
            selected_pf = st.selectbox("Select PF Number", pf_list, key="exit_pf")
            exit_date = st.date_input("Exit Date", date.today())
            manual_remark = st.text_input("Remark for Exit (optional)", key="exit_remark")
            if st.button("Mark Exited"):
                index = emp_df[emp_df["PF No."] == selected_pf].index
                emp_df.at[index, "Posting status"] = "EXITED"
                if manual_remark:
                    emp_df.at = manual_remark
                else:
                    emp_df.at = f"Transferred/Exited on {exit_date.strftime('%d-%m-%Y')}"
                employee_master["Apr.25"] = emp_df
                with pd.ExcelWriter("assets/EMPLOYEE MASTER DATA.xlsx", engine="openpyxl") as writer:
                    for sheet, df in employee_master.items():
                        df.to_excel(writer, sheet_name=sheet, index=False)
                st.success(f"Employee marked as exited.")

    # === Generate Document and Download (Refactored for In-Memory operation) ===
    if letter_type and template_files[letter_type] is not None:
        if st.button("Generate Document"):
            filename = ""
            if letter_type == "Exam NOC":
                filename = f"Exam_NOC_{context['PFNumber']}_{context}.docx"
            elif letter_type == "SF-11 Punishment Order":
                filename = f"SF-11_Punishment_{context['PFNumber']}_{context}.docx"
            elif letter_type == "Quarter Allotment Letter":
                filename = f"Quarter_Allotment_{context['PFNumber']}_{context['QuarterNo.']}.docx"
            elif letter_type == "DAR NOC Letter":
                filename = f"DAR_NOC_{context['PFNumber']}_{context}.docx"
            else:
                filename = f"{letter_type.replace(' ', '_')}_{context['PFNumber']}_{context}.docx"

            template_path = template_files[letter_type]
            
            # Generate the document in memory
            in_memory_file = generate_word(template_path, context)

            st.success(f"Document '{filename}' generated successfully!")
            
            # Create a download button for the in-memory file
            st.download_button(
                label="Download Document",
                data=in_memory_file,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
