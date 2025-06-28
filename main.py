import streamlit as st
import subprocess
import os

# Login system
users = {"admin": "Admin2123", "piyanshu": "PiY@2025_Secure!", "test": "test"}

if "user" not in st.session_state:
    st.title("🔐 Login")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    if st.button("Login"):
        if username in users and users[username] == password:
            st.session_state["user"] = username
            st.rerun()
        else:
            st.error("Invalid credentials")
    st.stop()

# After login
st.title("📊 Dashboard")
st.success(f"Logged in as: {st.session_state['user']}")

col1, col2 = st.columns(2)

with col1:
    if st.button("🧾 Open GST Invoice Extractor"):
        st.info("Launching excel_extract.py...")
        subprocess.Popen(["streamlit", "run", "excel_extrct.py"])

with col2:
    if st.button("📄 Open PDF to DOCX Filler"):
        st.info("Launching app.py...")
        subprocess.Popen(["streamlit", "run", "app.py"])

# Optional logout
if st.sidebar.button("🚪 Logout"):
    del st.session_state["user"]
    st.rerun()
