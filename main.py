import streamlit as st
import pandas as pd
import os
import base64
from docx import Document
from datetime import datetime, date, timedelta

# === Output Folder and Templates ===
os.makedirs("generated_letters", exist_ok=True)
template_files = {
    "Duty Letter (For Absent)": "assets/Absent Duty letter temp.docx",
    "SF-11 For Other Reason": "assets/SF-11 temp.docx",
    "Sick Memo": "assets/SICK MEMO temp..docx",
    "General Letter": "assets/General Letter temp.docx",
    "Exam NOC": "assets/Exam NOC Letter temp.docx",
    "SF-11 Punishment Order": "assets/SF-11 Punishment order temp.docx"
}

# === Load Data ===
employee_master = pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name=None)
sf11_register_path = "assets/SF-11 Register.xlsx"
sf11_register = pd.read_excel(sf11_register_path, sheet_name="SSE-SGAM")
noc_register_path = "assets/Exam NOC_Report.xlsx"

# === Placeholder Replace Function ===
from docx import Document

def replace_placeholder_in_para(paragraph, context):
    full_text = ''.join(run.text for run in paragraph.runs)
    replaced_text = full_text
    for key, val in context.items():
        replaced_text = replaced_text.replace(f"[{key}]", str(val))
    if full_text != replaced_text:
        # Clear all runs and set new single run
        for run in paragraph.runs:
            run.text = ''
        paragraph.runs[0].text = replaced_text

def generate_word(template_path, context, filename):
    doc = Document(template_path)
    # Replace in paragraphs
    for p in doc.paragraphs:
        replace_placeholder_in_para(p, context)
    # Replace in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_placeholder_in_para(p, context)
    save_path = os.path.join("generated_letters", filename)
    doc.save(save_path)
    return save_path

# === File Download ===
def download_word(path):
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    name = os.path.basename(path)
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{name}">📥 Download Word File</a>'
    st.markdown(href, unsafe_allow_html=True)

# === UI: Letter Type ===
st.title("📄 Letter Generator For OFFICE OF THE SSE/PW/SGAM")
letter_type = st.selectbox("📌 Select Letter Type:", list(template_files.keys()))

# === Employee Selection ===
sheet = st.selectbox("📋 Select Sheet", list(employee_master.keys()))
df = employee_master[sheet]
df["Display"] = df.apply(lambda r: f"{r[1]} - {r[2]} - {r[4]} - {r[5]}", axis=1)
selected = st.selectbox("👤 Select Employee", df["Display"].dropna())
row = df[df["Display"] == selected].iloc[0]

# === Extract Info ===
pf = row[1]
hrms = row[2]
unit_full = str(row[4])
unit = unit_full[:2]
station = row[8]
ename = row[5]
hname = row[13]
desg = row[18]
short = row[14]
letter_no = f"{short}/{unit}/{station}"

# === Common Context ===
letter_date = st.date_input("📅 Letter Date", value=date.today())
context = {
    "LetterDate": letter_date.strftime("%d-%m-%Y"),
    "EmployeeName": hname,
    "Designation": desg,
    "PFNumber": pf,
    "ShortName": short,
    "Unit": unit,
    "UnitNumber": unit,
    "LetterNo": letter_no,
    "DutyDate": "",
    "FromDate": "",
    "ToDate": "",
    "JoinDate": "",
    "Memo": ""
}

# === Letter Type Logic ===
if letter_type == "Duty Letter (For Absent)":
    st.subheader("🛠 Duty Letter")
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
    st.subheader("📌 SF-11 Other Reason")
    memo_input = st.text_area("Memo")
    context["Memo"] = memo_input + " जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है। अतः आप कामों व भूलो के फेहरिस्त धारा 1, 2 एवं 3 के उल्लंघन के दोषी पाए जाते है।"

elif letter_type == "Sick Memo":
    context["Memo"] = st.text_area("Memo")
    context["JoinDate"] = jd.strftime("%d-%m-%Y")

elif letter_type == "General Letter":
    officer = st.selectbox("TO",["सहायक मण्‍डल अभियंता प.म.रे. ब्‍योहारी",
                                 "मण्‍डल रेल प्रबंधक (कार्मिक) प.म.रे. जबलपुर मण्‍डल", 
                                 "मण्‍डल अभियंता (पूर्व) प.म.रे. जबलपुर मण्‍डल",
                                 "वरिष्‍ठ खण्‍ड अभियंता (रेल पथ) बरगवाँ",
                                 "वरिष्‍ठ खण्‍ड अभियंता (रेल पथ) सरईग्राम",
                                 "वरिष्‍ठ खण्‍ड अभियंता (रेल पथ) देवराग्राम",
                                 "कनिष्ठ अभियंता (रेल पथ) निवास रोड",
                                 "कनिष्ठ अभियंता (रेल पथ) भरसेड़ी",
                                 "कनिष्ठ अभियंता (रेल पथ) गजराबहरा",
                                 "कनिष्ठ अभियंता (रेल पथ) गोंदवाली"])
    subject = st.text_input("Subject")
    reference = st.text_input("Reference")
    memo_input = st.text_area("Detailed Memo")
    copies = st.selectbox("Copy To",["सहायक मण्‍डल अभियंता प.म.रे. ब्‍योहारी",
                                 "मण्‍डल रेल प्रबंधक (कार्मिक) प.म.रे. जबलपुर मण्‍डल", 
                                 "मण्‍डल अभियंता (पूर्व) प.म.रे. जबलपुर मण्‍डल",
                                 "वरिष्‍ठ खण्‍ड अभियंता (रेल पथ) बरगवाँ",
                                 "वरिष्‍ठ खण्‍ड अभियंता (रेल पथ) सरईग्राम",
                                 "वरिष्‍ठ खण्‍ड अभियंता (रेल पथ) देवराग्राम",
                                 "कनिष्ठ अभियंता (रेल पथ) निवास रोड",
                                 "कनिष्ठ अभियंता (रेल पथ) भरसेड़ी",
                                 "कनिष्ठ अभियंता (रेल पथ) गजराबहरा",
                                 "कनिष्ठ अभियंता (रेल पथ) गोंदवाली"])
    context["AddressTo"] = officer
    context["Subject"] = subject
    context["Reference"] = reference
    context["DetailMemo"] = memo_input
    context["CopyTo"] = "\n".join([c.strip() for c in copies.split(",")])

elif letter_type == "Exam NOC":
    exam_name = st.text_input("Exam Name")
    year = st.selectbox("NOC Year", [2025, 2024])
    df_noc = pd.read_excel(noc_register_path)
    count = sum((df_noc["PFNumber"] == pf) & (df_noc["Year"] == year))
    if count >= 4:
        st.warning("⚠️ Already 4 NOCs taken.")
    else:
        context["Memo"] = f"उपरोक्त कर्मचारी {exam_name} परीक्षा हेतु NOC हेतु पात्र है। यह इस वर्ष की {count+1}वीं स्वीकृति होगी।"

elif letter_type == "SF-11 Punishment Order":
    punishment = st.selectbox("Punishment Type", ["आगामी देय एक वर्ष की वेतन वृद्धि असंचयी प्रभाव से रोके जाने के अर्थदंड से दंडित किया जाता है।",
                                                  "आगामी देय एक वर्ष की वेतन वृद्धि संचयी प्रभाव से रोके जाने के अर्थदंड से दंडित किया जाता है।",
                                                  "आगामी देय एक सेट सुविधा पास तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।",
                                                  "आगामी देय एक सेट PTO तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।",
                                                  "आगामी देय दो सेट सुविधा पास तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।",
                                                  "आगामी देय दो सेट PTO तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।"])
    context["Memo"] = f"{punishment}"

# === GENERATE LETTER ===
if st.button("📄 Generate Letter"):
    temp = template_files[letter_type]
    fname = f"{letter_type.replace('/', '-')}-{hname}.docx"
    fpath = generate_word(temp, context, fname)
    st.success("✅ Letter generated!")
    download_word(fpath)

    # SF-11 Register Entry
    if letter_type in ["SF-11 For Other Reason", "SF-11 Punishment Order"] or (letter_type == "Duty Letter (For Absent)" and mode == "SF-11 & Duty Letter Only"):
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
   
    # Exam NOC Register Entry
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
