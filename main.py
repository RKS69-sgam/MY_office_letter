# SF-11 Punishment Order Logic
import streamlit as st
import pandas as pd
from datetime import date
from docx import Document
import os
from docx2pdf import convert
import base64

# Load SF-11 Register Data
sf11_df = pd.read_excel("assets/SF-11 Register.xlsx", sheet_name="SSE-SGAM")

# Employee Dropdown
sf11_df["Display"] = sf11_df.apply(
    lambda row: f"{row['पी.एफ. क्रमांक']} - {row['कर्मचारी का नाम']} - {row['दिनांक']} - {row['पत्र क्र.']}", axis=1)
selected_emp = st.selectbox("Select Employee (SF-11 Register):", sf11_df["Display"].tolist())
emp_row = sf11_df[sf11_df["Display"] == selected_emp].iloc[0]

# Letter Date
letter_date = st.date_input("📅 Letter Date", date.today())

# Dropdowns and Text Inputs
reply_received = st.selectbox("क्या कर्मचारी से प्रत्युत्तर प्राप्त हुआ?", ["हाँ", "नहीं"])
punishment_options = [
    "आगामी देय एक वर्ष की वेतन वृद्धि असंचयी प्रभाव से रोके जाने के अर्थदंड से दंडित किया जाता है।",
    "आगामी देय एक वर्ष की वेतन वृद्धि संचयी प्रभाव से रोके जाने के अर्थदंड से दंडित किया जाता है।",
    "आगामी देय एक सेट सुविधा पास तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।",
    "आगामी देय एक सेट PTO तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।",
    "आगामी देय दो सेट सुविधा पास तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।",
    "आगामी देय दो सेट PTO तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।"
]
punishment_text = st.selectbox("दंड का विवरण चुनें:", punishment_options)
order_date = st.date_input("दण्‍डादेश जारी करने का दिनांक")
appeal_date = st.date_input("यदि अपील की गई हो, तो अपील का दिनांक", value=None)
appeal_memo = st.text_area("अपील निर्णय पत्र क्र. एवं संक्षिप्त विवरण", "")
remarks = st.text_area("रिमार्क (यदि कोई हो)", "")

# Prepare Context
context = {
    "LetterDate": letter_date.strftime("%d-%m-%Y"),
    "Name": emp_row["कर्मचारी का नाम"],
    "Designation": emp_row["पदनाम"],
    "Memo": punishment_text,
    "PFNumber": emp_row["पी.एफ. क्रमांक"],
    "LetterNo": emp_row["पत्र क्र."],
    "PunishmentDate": order_date.strftime("%d-%m-%Y"),
    "AppealDate": appeal_date.strftime("%d-%m-%Y") if appeal_date else "",
    "AppealMemo": appeal_memo,
    "Remarks": remarks,
    "ReplyStatus": reply_received,
    "OrderNo": f"D-1/{emp_row['पत्र क्र.']}"
}

# Generate Document
def replace_placeholders(doc, ctx):
    for p in doc.paragraphs:
        for key, val in ctx.items():
            if f"[{key}]" in p.text:
                p.text = p.text.replace(f"[{key}]", str(val))
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, val in ctx.items():
                    if f"[{key}]" in cell.text:
                        cell.text = cell.text.replace(f"[{key}]", str(val))

if st.button("📄 Generate SF-11 Punishment Order"):
    template_path = "assets/SF-11 Punishment order temp.docx"
    doc = Document(template_path)
    replace_placeholders(doc, context)
    
    filename = f"SF11_Punishment_{context['PFNumber']}_{context['LetterDate']}"
    output_path = f"/tmp/{filename}.docx"
    doc.save(output_path)

    # Save to register (optional here)
    st.success("✅ Document Generated Successfully!")
    with open(output_path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
        href = f'<a href="data:application/octet-stream;base64,{b64}" download="{filename}.docx">⬇️ Download Word File</a>'
        st.markdown(href, unsafe_allow_html=True)

    # PDF conversion
    try:
        pdf_path = output_path.replace(".docx", ".pdf")
        convert(output_path, pdf_path)
        with open(pdf_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
            href = f'<a href="data:application/octet-stream;base64,{b64}" download="{filename}.pdf">⬇️ Download PDF File</a>'
            st.markdown(href, unsafe_allow_html=True)
    except:
        st.warning("⚠️ PDF conversion failed or not supported.")
