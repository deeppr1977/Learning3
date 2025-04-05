# Set the app title
import streamlit as st

st.set_page_config(page_title="📊 Learning Analytics Dashboard", layout="wide")

# Create main tab selector
selected_tab = st.sidebar.radio("📌 Choose View", ["Learning Insights Arena", "Build your own dashboard"])

# ---------------------------------------------------------------------------------------------------
# 🧠 SHEET 1: Learning Insights Arena (placeholder — insert your Sheet 1 logic here)
# ---------------------------------------------------------------------------------------------------
if selected_tab == "Learning Insights Arena":
    with open("Insights-Agent30.py", encoding="utf-8") as f:
        exec(f.read())

# --------------------------------------
# 📊 SHEET 2 – "Build your own dashboard"
# --------------------------------------
elif selected_tab == "Build your own dashboard":
    with open("Second_Sheet_Agent_2_6.py", encoding="utf-8") as f:
        exec(f.read())
