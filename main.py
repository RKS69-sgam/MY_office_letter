import streamlit as st

# Page Title
st.title("📄 Office Letter Generator")

# Step 1: Letter Type Selection
letter_type = st.selectbox(
    "🗂️ Select Letter Type:",
    [
        "SF-11 Punishment Order",
        "SF-11 For Other Reason",
        "Duty Letter (For Absent)",
        "Sick Memo",
        "Exam NOC",
        "General Letter"
    ]
)

st.markdown(f"✅ Selected Letter Type: **{letter_type}**")