# FINAL STREAMLIT LETTER GENERATOR APP WITH ALL FIXES

import streamlit as st
import pandas as pd
import os
import base64
from docx import Document
from datetime import date, timedelta

# Create output folder
os.makedirs("generated_letters", exist_ok=True)

# File paths
template_files = {
    "Duty Letter (For Absent)": "assets/Absent Duty letter temp.docx",
    "SF-11 For Other Reason": "assets/SF-11 temp.docx",
    "Sick Memo": "assets/SICK MEMO temp..docx",
    "General Letter": "assets/General Letter temp.docx",
    "Exam NOC": "assets/Exam NOC Letter temp.docx",
    "SF-11 Punishment Order": "assets/SF-11 Punishment order temp.docx"
}

employee_master = pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name=None)
sf11_register_path = "assets/SF-11 Register.xlsx"
sf11_register = pd.read_excel(sf11_register_path, sheet_name="SSE-SGAM")
noc_register_path = "assets/Exam NOC_Report.xlsx"
df_noc = pd.read_excel(noc_register_path) if os.path.exists(noc_register_path) else pd.DataFrame(columns=["PF Number", "Employee Name", "Designation", "NOC Year", "Application No.", "Exam Name"])

# Replace function for paragraphs and tables
def replace_placeholder_in_para(paragraph, context):
    full_text = ''.join(run.text for run in paragraph.runs)
    new_text = full_text
    for key, val in context.items():
        new_text = new_text.replace(f"[{key}]", str(val))
    if new_text != full_text:
        for run in paragraph.runs:
            run.text = ''
        if paragraph.runs:
            paragraph.runs[0].text = new_text
        else:
            paragraph.add_run(new_text)

def generate_word(template_path, context, filename):
    doc = Document(template_path)
    for p in doc.paragraphs:
        replace_placeholder_in_para(p, context)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_placeholder_in_para(p, context)
    output_path = os.path.join("generated_letters", filename)
    doc.save(output_path)
    return output_path

def download_word(path):
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    name = os.path.basename(path)
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{name}">Download Word File</a>'
    st.markdown(href, unsafe_allow_html=True)

# === UI ===
st.title("OFFICE OF THE SSE/PW/SGAM")
letter_type = st.selectbox("Select Letter Type", list(template_files.keys()))

# === Select Employee Logic ===
if letter_type == "SF-11 Punishment Order":
    df = sf11_register
    df["Display"] = df.apply(lambda r: f"{r['पी.एफ. क्रमांक']} - {r['कर्मचारी का नाम']} - {r['पत्र क्र.']}", axis=1)
    selected = st.selectbox("Select Employee", df["Display"].dropna())
    row = df[df["Display"] == selected].iloc[0]
    patra_kr = row["पत्र क्र."]
    dandadesh_krmank = f"{patra_kr}/D-1"
    pf = row["पी.एफ. क्रमांक"]
    hname = row["कर्मचारी का नाम"]
    desg = row.get("पदनाम", "")
    unit_full = patra_kr.split("/")[1]
    unit = unit_full[:2]
    short = patra_kr.split("/")[0]
    letter_no = dandadesh_krmank
elif letter_type == "General Letter":
    df = pd.DataFrame()
    pf = hname = desg = unit = unit_full = short = letter_no = ""
else:
    df = employee_master["Apr.25"]
    df["Display"] = df.apply(lambda r: f"{r[1]} - {r[2]} - {r[4]} - {r[5]}", axis=1)
    selected = st.selectbox("Select Employee", df["Display"].dropna())
    row = df[df["Display"] == selected].iloc[0]
    pf = row[1]
    hname = row[13]
    desg = row[18]
    unit_full = str(row[4])
    unit = unit_full[:2]
    short = row[14]
    letter_no = f"{short}/{unit}/{unit_full}"

letter_date = st.date_input("Letter Date", value=date.today())
context = {
    "LetterDate": letter_date.strftime("%d-%m-%Y"),
    "EmployeeName": hname,
    "Designation": desg,
    "PFNumber": pf,
    "ShortName": short,
    "Unit": unit,
    "UnitNumber": unit_full,
    "LetterNo": letter_no,
    "DutyDate": "",
    "FromDate": "",
    "ToDate": "",
    "JoinDate": "",
    "Memo": "",
    "OfficerUnit": "",
    "Subject": "",
    "Reference": "",
    "CopyTo": ""
}

if letter_type == "Duty Letter (For Absent)":
    mode = st.selectbox("Mode", ["SF-11 & Duty Letter Only", "Duty Letter Only"])
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
    context["FileName"] = st.selectbox("File Name", [
        "", "STAFF-IV", "OFFICE ORDER", "STAFF-III", "QAURTER-1", "ARREAR",
        "CEA/STAFF-IV", "CEA/STAFF-III", "PW-SGAM", "MISC."
    ])

    officer_option = st.selectbox("अधिकारी/कर्मचारी", [
        "", "सहायक मण्‍डल अभियंता", "मण्‍डल अभिंयता (पूर्व)", "मण्‍डल अभिंयता (पश्चिम)",
        "मण्‍डल रेल प्रबंधक (कार्मिक)", "मण्‍डल रेल प्रबंधक (कार्य)", "वरिष्‍ठ खण्‍ड अभियंता (रेल पथ)",
        "वरिष्‍ठ खण्‍ड अभियंता (कार्य)", "वरिष्‍ठ खण्‍ड अभियंता (विद्युत)", "वरिष्‍ठ खण्‍ड अभियंता (T&D)",
        "वरिष्‍ठ खण्‍ड अभियंता (S&T)", "वरिष्‍ठ खण्‍ड अभियंता (USFD)", "वरिष्‍ठ खण्‍ड अभियंता (PW/STORE)",
        "कनिष्‍ठ अभियंता (रेल पथ)", "कनिष्‍ठ अभियंता (कार्य)", "कनिष्‍ठ अभियंता (विद्युत)",
        "कनिष्‍ठ अभियंता (T&D)", "कनिष्‍ठ अभियंता (S&T)", "शाखा सचिव (WCRMS)",
        "मण्‍डल अध्‍यक्ष (WCRMS)", "मण्‍डल सचिव (WCRMS)", "महामंत्री (WCRMS)", "अन्‍य"
    ])
    if officer_option == "अन्‍य":
        officer_option = st.text_input("अन्‍य का नाम/पदनाम/एजेंसी का नाम लिखें")
    context["OfficerName"] = officer_option

    address_option = st.selectbox("पता", [
        "", "प.म.रे. ब्‍योहारी", "प.म.रे. जबलपुर", "सरईग्राम", "देवराग्राम", "बरगवॉं",
        "निवासरोड", "भरसेड़ी", "गजराबहरा", "गोंदवाली", "अन्‍य"
    ])
    if address_option == "अन्‍य":
        address_option = st.text_input("अन्‍य का पता लिखें")
    context["OfficeAddress"] = address_option

    subject_input = st.text_input("विषय")
    context["Subject"] = f"विषय:-    {subject_input}" if subject_input.strip() else ""

    ref_input = st.text_input("संदर्भ")
    context["Reference"] = f"संदर्भ:-    {ref_input}" if ref_input.strip() else ""

    context["Memo"] = st.text_area("मुख्‍य विवरण")

    copy_input = st.text_input("प्रतिलिपि")
    context["CopyTo"] = f"प्रतिलिपि:-    " + "\n".join([c.strip() for c in copy_input.split(",") if c.strip()]) if copy_input.strip() else ""

elif letter_type == "Exam NOC":
    exam_name = st.text_input("Exam Name")
    year = st.selectbox("NOC Year", [2025, 2024])
    count = sum((df_noc["PF Number"] == pf) & (df_noc["NOC Year"] == year))
    if count >= 4:
        st.warning("Already 4 NOCs taken.")
    else:
        application_no = count + 1
        table_text = f"""| PF Number | Employee Name | Designation | NOC Year | Application No. | Exam Name |
|------------|----------------|-------------|----------|----------------|------------|
| {pf} | {hname} | {desg} | {year} | {application_no} | {exam_name} |
"""
        context["PFNumber"] = table_text
elif letter_type == "SF-11 Punishment Order":
    context["Memo"] = st.selectbox("Punishment Type", [
        "आगामी देय एक वर्ष की वेतन वृद्धि असंचयी प्रभाव से रोके जाने के अर्थदंड से दंडित किया जाता है।",
        "आगामी देय एक वर्ष की वेतन वृद्धि संचयी प्रभाव से रोके जाने के अर्थदंड से दंडित किया जाता है।",
        "आगामी देय एक सेट सुविधा पास तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।",
        "आगामी देय एक सेट PTO तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।",
        "आगामी देय दो सेट सुविधा पास तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।",
        "आगामी देय दो सेट PTO तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।"
    ])

if st.button("Generate Letter"):
    if letter_type == "Duty Letter (For Absent)" and mode == "SF-11 & Duty Letter Only":
        duty_path = generate_word(template_files["Duty Letter (For Absent)"], context, f"DutyLetter-{hname}.docx")
        sf11_path = generate_word(template_files["SF-11 For Other Reason"], context, f"SF-11-{hname}.docx")
        download_word(duty_path)
        download_word(sf11_path)
    else:
        word_path = generate_word(template_files[letter_type], context, f"{letter_type.replace('/', '-')}-{hname}.docx")
        download_word(word_path)

    if letter_type in ["SF-11 For Other Reason", "Duty Letter (For Absent)"]:
        new_entry = pd.DataFrame([{ 
            "पी.एफ. क्रमांक": pf,
            "कर्मचारी का नाम": hname,
            "पदनाम": desg,
            "पत्र क्र.": letter_no,
            "दिनांक": letter_date.strftime("%d-%m-%Y"),
            "दण्ड का विवरण": context["Memo"]
        }])
        sf11_register = pd.concat([sf11_register, new_entry], ignore_index=True)
        sf11_register.to_excel(sf11_register_path, sheet_name="SSE-SGAM", index=False)

    if letter_type == "SF-11 Punishment Order":
        mask = (sf11_register["पी.एफ. क्रमांक"] == pf) & (sf11_register["पत्र क्र."] == patra_kr)
        if mask.any():
            i = sf11_register[mask].index[0]
            sf11_register.at[i, "दण्डादेश क्रमांक"] = letter_no
            sf11_register.at[i, "दण्ड का विवरण"] = context["Memo"]
            sf11_register.to_excel(sf11_register_path, sheet_name="SSE-SGAM", index=False)
        else:
            st.warning("\u26a0\ufe0f चयनित कर्मचारी के लिए पत्र क्र. के आधार पर प्रविष्टि नहीं मिली।")

    if letter_type == "Exam NOC" and count < 4:
        new_noc = {
            "PF Number": pf,
            "Employee Name": hname,
            "Designation": desg,
            "NOC Year": year,
            "Application No.": count + 1,
            "Exam Name": exam_name
        }
        df_noc = pd.concat([df_noc, pd.DataFrame([new_noc])], ignore_index=True)
        df_noc.to_excel(noc_register_path, index=False)
