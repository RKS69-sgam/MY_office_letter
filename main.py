import streamlit as st
import pandas as pd
import os
import base64
from docx import Document
from datetime import datetime, date, timedelta

# === OUTPUT FOLDERS & TEMPLATE PATHS ===
os.makedirs("generated_letters", exist_ok=True)
template_files = {
    "Duty Letter (For Absent)": "assets/Absent Duty letter temp.docx",
    "SF-11 For Other Reason": "assets/SF-11 temp.docx",
    "Sick Memo": "assets/SICK MEMO temp..docx",
    "General Letter": "assets/General Letter temp.docx",
    "Exam NOC": "assets/Exam NOC Letter temp.docx",
    "SF-11 Punishment Order": "assets/SF-11 Punishment order temp.docx"
}

# === LOAD MASTER DATA ===
employee_master = pd.read_excel("assets/EMPLOYEE MASTER DATA.xlsx", sheet_name=None)
sf11_register = pd.read_excel("assets/SF-11 Register.xlsx", sheet_name="SSE-SGAM")
noc_register_path = "assets/Exam NOC_Report.xlsx"

# === PLACEHOLDER REPLACEMENT FUNCTION ===
def generate_word(template_path, context, filename):
    doc = Document(template_path)
    for p in doc.paragraphs:
        for key, val in context.items():
            if f"[{key}]" in p.text:
                inline = p.runs
                for i in range(len(inline)):
                    if f"[{key}]" in inline[i].text:
                        inline[i].text = inline[i].text.replace(f"[{key}]", str(val))
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, val in context.items():
                        if f"[{key}]" in p.text:
                            inline = p.runs
                            for i in range(len(inline)):
                                if f"[{key}]" in inline[i].text:
                                    inline[i].text = inline[i].text.replace(f"[{key}]", str(val))
    save_path = os.path.join("generated_letters", filename)
    doc.save(save_path)
    return save_path

# === DOWNLOAD LINK ===
def download_word(path):
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    name = os.path.basename(path)
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{name}">📥 Download Word File</a>'
    st.markdown(href, unsafe_allow_html=True)

# === UI: LETTER TYPE SELECTION ===
st.title("📄 Railway Letter Generator")
letter_type = st.selectbox("📌 Select Letter Type:", list(template_files.keys()))

# === SELECT EMPLOYEE ===
sheet = st.selectbox("📋 Select Sheet", list(employee_master.keys()))
df = employee_master[sheet]
df["Display"] = df.apply(lambda r: f"{r[1]} - {r[2]} - {r[4]} - {r[5]}", axis=1)
selected = st.selectbox("👤 Select Employee", df["Display"].dropna())
row = df[df["Display"] == selected].iloc[0]

# === EXTRACT FIELDS ===
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

# === DATES & CONTEXT ===
letter_date = st.date_input("📅 Letter Date", value=date.today())
context = {
    "LetterDate": letter_date.strftime("%d-%m-%Y"),
    "EmployeeName": hname,
    "Designation": desg,
    "PFNumber": pf,
    "ShortName": short,
    "Unit": unit,
    "UnitNumber": unit_full,
    "LetterNo": letter_no,
    "DutyDate": "",  # for Duty Letter
    "FromDate": "",
    "ToDate": "",
    "JoinDate": "",
    "Memo": ""
}

# === INDIVIDUAL CONDITIONS ===
if letter_type == "Duty Letter (For Absent)":
    st.subheader("🛠 Duty Letter")
    mode = st.selectbox("Mode", ["SF-11 & Duty Letter For Absent", "Duty Letter For Absent"])
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
    context["Memo"] = memo_input + " जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है।"

elif letter_type == "Sick Memo":
    context["Memo"] = st.text_area("Memo")
    jd = st.date_input("Join Date", value=date.today())
    context["JoinDate"] = jd.strftime("%d-%m-%Y")

elif letter_type == "General Letter":
    officer = st.text_input("Officer / Unit to whom addressed")
    subject = st.text_input("Subject")
    reference = st.text_input("Reference")
    memo_input = st.text_area("Detailed Memo")
    copies = st.text_area("Copy to (comma-separated)")

    context["OfficerUnit"] = officer
    context["Subject"] = subject
    context["Reference"] = reference
    context["Memo"] = memo_input
    context["CopyTo"] = "\n".join([c.strip() for c in copies.split(",")])

elif letter_type == "Exam NOC":
    exam_name = st.text_input("Exam Name")
    year = st.selectbox("NOC Year", [2025, 2024])
    count = sum((pd.read_excel(noc_register_path)["PFNumber"] == pf) & 
                (pd.read_excel(noc_register_path)["Year"] == year))

    if count >= 4:
        st.warning("⚠️ Already 4 NOCs taken in selected year.")
    else:
        context["Memo"] = f"उपरोक्त कर्मचारी {exam_name} परीक्षा हेतु NOC हेतु पात्र है। यह इस वर्ष की {count+1}वीं स्वीकृति होगी।"

elif letter_type == "SF-11 Punishment Order":
    punishment = st.selectbox("Punishment Type", ["Censure", "Withholding", "Reduction"])
    context["Memo"] = f"आपके द्वारा की गई अनुशासनहीनता के लिए यह दंड प्रदान किया जाता है: {punishment}"

# === FINAL GENERATION ===
if st.button("📄 Generate Letter"):
    temp = template_files[letter_type]
    fname = f"{letter_type.replace('/', '-')}-{hname}.docx"
    fpath = generate_word(temp, context, fname)
    st.success("✅ Letter generated successfully!")
    download_word(fpath)

    # === REGISTERS ENTRY LOGIC ===
    if letter_type in ["SF-11 For Other Reason", "SF-11 Punishment Order"]:
        new_entry = pd.DataFrame([{
            "PFNumber": pf,
            "Name": hname,
            "Designation": desg,
            "Letter No.": letter_no,
            "Letter Date": letter_date.strftime("%d-%m-%Y"),
            "Memo": context["Memo"]
        }])
        updated = pd.concat([sf11_register, new_entry], ignore_index=True)
        updated.to_excel("assets/SF-11 Register.xlsx", sheet_name="SSE-SGAM", index=False)

    if letter_type == "Exam NOC" and count < 4:
        df_noc = pd.read_excel(noc_register_path)
        new_row = {
            "PFNumber": pf,
            "Name": hname,
            "Year": year,
            "Exam": exam_name,
            "Date": letter_date.strftime("%d-%m-%Y"),
            "Memo": context["Memo"]
        }
        df_noc = df_noc.append(new_row, ignore_index=True)
        df_noc.to_excel(noc_register_path, index=False)
