import streamlit as st
import pandas as pd
import os
import base64
from docx import Document
from datetime import date, timedelta

# Create output folder
os.makedirs("generated_letters", exist_ok=True)

template_files = {
    "Duty Letter (For Absent)": "assets/Absent Duty letter temp.docx",
    "SF-11 For Other Reason": "assets/SF-11 temp.docx",
    "Sick Memo": "assets/SICK MEMO temp..docx",
    "General Letter": "assets/General Letter temp.docx",
    "Exam NOC": "assets/Exam NOC Letter temp.docx",
    "SF-11 Punishment Order": "assets/SF-11 Punishment order temp.docx"
}

# Load data
employee_master = pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name=None)
sf11_register_path = "assets/SF-11 Register.xlsx"
sf11_register = pd.read_excel(sf11_register_path, sheet_name="SSE-SGAM")
noc_register_path = "assets/Exam NOC_Report.xlsx"
df_noc = pd.read_excel(noc_register_path) if os.path.exists(noc_register_path) else pd.DataFrame(columns=["PFNumber", "Name", "Year", "Exam", "Date", "Memo"])

# Word replace
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

# UI
st.title("OFFICE OF THE SSE/PW/SGAM")
letter_type = st.selectbox("Select Letter Type", list(template_files.keys()))

# === Employee Data Source ===
if letter_type == "SF-11 Punishment Order":
    df = sf11_register
    df["Display"] = df.apply(lambda r: f"{r['पी.एफ. क्रमांक']} - {r['कर्मचारी का नाम']} - {r['पत्र क्र.']}", axis=1)
    patra_kr = row["पत्र क्र."]  # Already filled in the register
    dandadesh_krmank = f"{patra_kr}/D-1"
elif letter_type == "General Letter":
    df = pd.DataFrame()
else:
    df = employee_master["Apr.25"]
    df["Display"] = df.apply(lambda r: f"{r[1]} - {r[2]} - {r[4]} - {r[5]}", axis=1)
elif letter_type == "General Letter":
    df = pd.DataFrame()
else:
    df = employee_master["Apr.25"]
    df["Display"] = df.apply(lambda r: f"{r[1]} - {r[2]} - {r[4]} - {r[5]}", axis=1)

if letter_type != "General Letter":
    selected = st.selectbox("Select Employee", df["Display"].dropna())
    row = df[df["Display"] == selected].iloc[0]
    pf = row[1]
    hname = row[13] if letter_type != "SF-11 Punishment Order" else row["Name"]
    desg = row[18] if letter_type != "SF-11 Punishment Order" else row["Designation"]
    unit_full = str(row[4]) if letter_type != "SF-11 Punishment Order" else row["Letter No."].split("/")[1]
    unit = unit_full[:2]
    short = row[14] if letter_type != "SF-11 Punishment Order" else row["Letter No."].split("/")[0]
    letter_no = f"{short}/{unit}/{unit_full}"
else:
    pf, hname, desg, unit, unit_full, short, letter_no = "", "", "", "", "", "", ""

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
    context["FromDate"] = fd.strftime("%d-%m-%Y")
    context["ToDate"] = td.strftime("%d-%m-%Y")
    context["JoinDate"] = jd.strftime("%d-%m-%Y")
    context["DutyDate"] = jd.strftime("%d-%m-%Y")
    days = (td - fd).days + 1
    context["Memo"] = f"आप बिना किसी पूर्व सूचना के दिनांक {fd.strftime('%d-%m-%Y')} से {td.strftime('%d-%m-%Y')} तक कुल {days} दिवस कार्य से अनुपस्थित थे, जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है। अतः आप कामों व भूलो के फेहरिस्त धारा 1, 2 एवं 3 के उल्लंघन के दोषी पाए जाते है।"

elif letter_type == "SF-11 For Other Reason":
    memo_input = st.text_area("Memo")
    context["Memo"] = memo_input + " जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है। अतः आप कामों व भूलो के फेहरिस्त धारा 1, 2 एवं 3 के उल्लंघन के दोषी पाए जाते है।"

#elif letter_type == "Sick Memo":
    #memo_input = st.text_area("Memo")
   # context["Memo"] = memo_input

elif letter_type == "General Letter":
    officer = st.text_area("To Officer/Unit")
    subject = st.text_input("Subject")
    reference = st.text_input("Reference")
    memo_input = st.text_area("Detailed Memo")
    copies =  st.text_input("Copy To (comma-separated)")
    context["AddressTo"] = officer
    context["Subject"] = subject
    context["Reference"] = reference
    context["DetailMemo"] = memo_input
    context["CopyTo"] = "\n".join([c.strip() for c in copies.split(",")])

elif letter_type == "Exam NOC":
    exam_name = st.text_input("Exam Name")
    year = st.selectbox("NOC Year", [2025, 2024])
    count = sum((df_noc["PFNumber"] == pf) & (df_noc["Year"] == year))
    if count >= 4:
        st.warning("Already 4 NOCs taken.")
    else:
        context["Memo"] = f"उपरोक्त कर्मचारी {exam_name} परीक्षा हेतु NOC हेतु पात्र है। यह इस वर्ष की {count+1}वीं स्वीकृति होगी।"

elif letter_type == "SF-11 Punishment Order":
    punishment = st.selectbox("Punishment Type", [
        "आगामी देय एक वर्ष की वेतन वृद्धि असंचयी प्रभाव से रोके जाने के अर्थदंड से दंडित किया जाता है।",
        "आगामी देय एक वर्ष की वेतन वृद्धि संचयी प्रभाव से रोके जाने के अर्थदंड से दंडित किया जाता है।",
        "आगामी देय एक सेट सुविधा पास तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।",
        "आगामी देय एक सेट PTO तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।",
        "आगामी देय दो सेट सुविधा पास तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।",
        "आगामी देय दो सेट PTO तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।"
    ])
    context["Memo"] = punishment

if st.button("Generate Letter"):
    if letter_type == "Duty Letter (For Absent)" and mode == "SF-11 & Duty Letter Only":
        # === Duty Letter ===
        duty_template = template_files["Duty Letter (For Absent)"]
        duty_filename = f"DutyLetter-{hname}.docx"
        duty_path = generate_word(duty_template, context, duty_filename)

        # === SF-11 Letter ===
        sf11_template = template_files["SF-11 For Other Reason"]
        sf11_filename = f"SF-11-{hname}.docx"
        sf11_path = generate_word(sf11_template, context, sf11_filename)

        # ✅ Show both download links
        st.success("✅ SF-11 & Duty Letter Generated")
        download_word(duty_path)
        download_word(sf11_path)

        # ✅ Entry in SF-11 Register
        new_entry = pd.DataFrame([{
            "PFNumber": pf,
            "Name": hname,
            "Designation": desg,
            "Letter No.": letter_no,
            "Letter Date": letter_date.strftime("%d-%m-%Y"),
            "Memo": context["Memo"]
        }])
        updated = pd.concat([sf11_register, new_entry], ignore_index=True)
        updated.to_excel(sf11_register_path, sheet_name="SSE-SGAM", index=False)

    else:
        # === For all other cases: generate one letter ===
        template = template_files[letter_type]
        filename = f"{letter_type.replace('/', '-')}-{hname}.docx"
        path = generate_word(template, context, filename)
        st.success("✅ Letter generated successfully!")
        download_word(path)

        # === SF-11 Register entry for relevant types ===

if (
    letter_type == "SF-11 For Other Reason"
    or (letter_type == "Duty Letter (For Absent)" and mode == "SF-11 & Duty Letter Only")
):
    new_entry = pd.DataFrame([{
        "पी.एफ. क्रमांक": pf,
        "कर्मचारी का नाम": hname,
        "पदनाम": desg,
        "पत्र क्र.": letter_no,
        "दिनांक": letter_date.strftime("%d-%m-%Y"),
        "दण्ड का विवरण": context["Memo"],
    }])
    updated = pd.concat([sf11_register, new_entry], ignore_index=True)
    updated.to_excel(sf11_register_path, sheet_name="SSE-SGAM", index=False)

if letter_type == "SF-11 Punishment Order":
    # Find matching row in SF-11 Register by PFNumber and Letter No.
    mask = (sf11_register["पी.एफ. क्रमांक"] == pf) & (sf11_register["पत्र क्र."] == Patra_kr)
    if mask.any():
        idx = sf11_register[mask].index[0]
        sf11_register.at[idx, "दण्डादेश क्रमांक"] = dandadesh_krmank
        sf11_register.at[idx, "दण्ड का विवरण"] = context["Memo"]
    else:
        st.warning("⚠️ चयनित कर्मचारी के लिए पत्र क्र. के आधार पर प्रविष्टि नहीं मिली। कृपया SF-11 Register जांचें।")
    # Save updated file
    sf11_register.to_excel(sf11_register_path, sheet_name="SSE-SGAM", index=False)
        

        # === Exam NOC Register entry ===
        if letter_type == "Exam NOC" and count < 4:
            new_noc = {
                "PFNumber": pf,
                "Name": hname,
                "Year": year,
                "Exam": exam_name,
                "Date": letter_date.strftime("%d-%m-%Y"),
                "Memo": context["Memo"]
            }
            df_noc = df_noc.append(new_noc, ignore_index=True)
            df_noc.to_excel(noc_register_path, index=False)

