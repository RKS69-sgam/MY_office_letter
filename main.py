import streamlit as st
from datetime import date, timedelta

# === App Title ===
st.title("📄 Railway Letter Generator")

# === Step 1: Select Letter Type ===
letter_types = [
    "Duty Letter (For Absent)",
    "SF-11 For Other Reason",
    "Sick Memo",
    "General Letter",
    "Exam NOC",
    "SF-11 Punishment Order"
]
selected_letter_type = st.selectbox("1️⃣ Select Letter Type:", letter_types)

# === Step 2: Duty Letter (For Absent) Logic ===
if selected_letter_type == "Duty Letter (For Absent)":
    st.subheader("📌 Duty Letter Details")

    # Select Duty Letter Type (combined or only duty)
    duty_mode = st.selectbox("✏️ Duty Letter Mode", [
        "SF-11 & Duty Letter For Absent",
        "Duty Letter For Absent"
    ])

    # Select from and to dates
    from_date = st.date_input("📅 From Date", value=date.today() - timedelta(days=3))
    to_date = st.date_input("📅 To Date", value=date.today())

    # Auto-set join date = next day of to_date
    suggested_join_date = to_date + timedelta(days=1)
    duty_join_date = st.date_input("📆 Expected Join Date", suggested_join_date)

    # Summary
    st.markdown("### 📄 Summary")
    st.write(f"**Mode:** {duty_mode}")
    st.write(f"**From:** {from_date.strftime('%d-%m-%Y')}")
    st.write(f"**To:** {to_date.strftime('%d-%m-%Y')}")
    st.write(f"**Join Date:** {duty_join_date.strftime('%d-%m-%Y')}")

    # Context for template
    context = {
        "FromDate": from_date.strftime("%d-%m-%Y"),
        "ToDate": to_date.strftime("%d-%m-%Y"),
        "JoinDate": duty_join_date.strftime("%d-%m-%Y"),
        "DutyMode": duty_mode
    }

    # Optional preview of context
    st.markdown("### 🔍 Template Context")
    st.json(context)