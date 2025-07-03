import streamlit as st
from datetime import date, timedelta

st.title("📌 Duty Letter (For Absent) Generator")

# Step 1: Select Duty Letter Type
duty_mode = st.selectbox("✏️ Select Duty Letter Type:", [
    "SF-11 & Duty Letter For Absent",
    "Duty Letter For Absent"
])

# Step 2: Select Date Range
from_date = st.date_input("📅 From Date", value=date.today() - timedelta(days=3))
to_date = st.date_input("📅 To Date", value=date.today())

# Step 3: Auto-Suggest Join Date (Next day after To Date)
suggested_join_date = to_date + timedelta(days=1)
duty_join_date = st.date_input("📆 Expected Join Date", suggested_join_date)

# Step 4: Show Summary
st.markdown("### 📄 Summary")
st.write(f"**Duty Mode:** {duty_mode}")
st.write(f"**From Date:** {from_date.strftime('%d-%m-%Y')}")
st.write(f"**To Date:** {to_date.strftime('%d-%m-%Y')}")
st.write(f"**Expected Join Date:** {duty_join_date.strftime('%d-%m-%Y')}")

# Step 5: Create Context Dictionary for Template
duty_context = {
    "FromDate": from_date.strftime("%d-%m-%Y"),
    "ToDate": to_date.strftime("%d-%m-%Y"),
    "JoinDate": duty_join_date.strftime("%d-%m-%Y"),
    "DutyMode": duty_mode
}

# Step 6: Optional – Show Dictionary
st.markdown("### 🔍 Context Dictionary for Template")
st.json(duty_context)