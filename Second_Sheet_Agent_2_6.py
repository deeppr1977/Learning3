
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import os
import smtplib
from dotenv import load_dotenv
from datetime import datetime
from langchain.chat_models import ChatOpenAI
from langchain_experimental.agents import create_pandas_dataframe_agent
from fpdf import FPDF
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText

# Load environment variables
load_dotenv()
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

@st.cache_data
def load_data():
    df = pd.read_excel("Dummy_Course_Data.xlsx", engine="openpyxl")
    df["Course Completion Date"] = pd.to_datetime(df["Course Completion Date"], errors="coerce")
    df["Registration Date"] = pd.to_datetime(df["Registration Date"], errors="coerce")
    return df

df = load_data()

def gpt_suggest_insights(prompt_text):
    llm = ChatOpenAI(openai_api_key=OPENAI_API_KEY, model_name="gpt-3.5-turbo", temperature=0)
    agent = create_pandas_dataframe_agent(llm, df, verbose=False, handle_parsing_errors=True, allow_dangerous_code=True)
    return agent.run(prompt_text)

def create_pdf_report_with_charts(insight_data, filename="final_report.pdf"):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    for metric, (insight, chart_path) in insight_data.items():
        pdf.set_font("Arial", 'B', 14)
        pdf.multi_cell(0, 10, f"Metric: {metric}")
        pdf.set_font("Arial", '', 12)
        pdf.multi_cell(0, 10, insight)
        if chart_path and os.path.exists(chart_path):
            pdf.image(chart_path, w=170)
        pdf.ln(10)

    pdf.output(filename)
    return filename

def send_email_with_pdf(to_email, subject, body, pdf_path):
    try:
        sender = os.getenv("EMAIL_USERNAME")
        password = os.getenv("EMAIL_PASSWORD")
        host = os.getenv("EMAIL_HOST")
        port = int(os.getenv("EMAIL_PORT"))

        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = to_email
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain"))

        with open(pdf_path, "rb") as f:
            part = MIMEApplication(f.read(), _subtype="pdf")
            part.add_header("Content-Disposition", f"attachment; filename=" + os.path.basename(pdf_path))
            msg.attach(part)

        with smtplib.SMTP(host, port) as server:
            server.starttls()
            server.login(sender, password)
            server.sendmail(sender, to_email, msg.as_string())
        return True
    except Exception as e:
        st.error(f"Email Error: {e}")
        return False

def render_chart(metric, chart_type, ax):
    if metric == "Completion by Platform":
        data = df["Platform"].value_counts()
    elif metric == "Completion by Employee Role":
        data = df["Employee Role"].value_counts()
    elif metric == "Completion by Organization":
        data = df["Main Organization Unit"].value_counts()
    elif metric == "Top 5 courses by completion":
        data = df["Course Name"].value_counts().nlargest(5)
    elif metric == "Bottom 5 courses by completion":
        data = df["Course Name"].value_counts().nsmallest(5)
    elif metric == "Currently Enrolled":
        data = df[df["Course Completion Date"].isna()]["Course Name"].value_counts().nlargest(5)
    elif metric == "Number of completions":
        data = df["Course Completion Date"].dt.month.value_counts().sort_index()
    elif metric == "Completion variance to previous month":
        month_counts = df["Course Completion Date"].dt.to_period("M").value_counts().sort_index()
        data = month_counts.diff().fillna(0)
    elif metric == "Number of employees registered vs completed (monthly trend)":
        reg = df["Registration Date"].dt.to_period("M").value_counts().sort_index()
        comp = df["Course Completion Date"].dt.to_period("M").value_counts().sort_index()
        ax.plot(reg.index.astype(str), reg.values, label="Registered")
        ax.plot(comp.index.astype(str), comp.values, label="Completed")
        ax.legend()
        return

    if chart_type == "Bar":
        data.plot(kind="bar", ax=ax)
    elif chart_type == "Line":
        data.plot(kind="line", ax=ax)
    elif chart_type == "Pie":
        data.plot(kind="pie", ax=ax, autopct="%1.1f%%")
    elif chart_type == "Table":
        ax.axis("off")
        table_data = pd.DataFrame({metric: data.index, "Value": data.values})
        ax.table(cellText=table_data.values, colLabels=table_data.columns, loc="center")

# Initialize session state
if "report_data" not in st.session_state:
    st.session_state["report_data"] = {}
if "final_pdf" not in st.session_state:
    st.session_state["final_pdf"] = None

st.title("📘 Sheet 2 – AI-Guided Insight Dashboard")

mode = st.radio("Choose input mode:", ["Let AI generate report", "Choose metrics manually"])

metrics = [
    "Currently Enrolled", "Number of completions", "Completion variance to previous month",
    "Number of employees registered vs completed (monthly trend)",
    "Top 5 courses by completion", "Bottom 5 courses by completion",
    "Completion by Platform", "Completion by Employee Role", "Completion by Organization"
]

if mode == "Let AI generate report":
    selected_metrics = metrics[:4]
else:
    selected_metrics = st.multiselect("Select metrics to explore:", metrics)

custom_text = st.text_area("📝 Add custom context (optional):")

if st.button("🔍 Generate Insights & Charts"):
    if not selected_metrics:
        st.warning("Please select metrics or switch to AI mode.")
    else:
        st.session_state["report_data"] = {}
        for metric in selected_metrics:
            insight = gpt_suggest_insights(f"Explain: {metric}. Context: {custom_text}. Keep it concise (5 lines).")
            st.session_state["report_data"][metric] = {"insight": insight, "chart_type": "Bar"}

if st.session_state["report_data"]:
    for metric in st.session_state["report_data"].keys():
        st.subheader(f"🧠 Insight for {metric}")
        st.markdown(st.session_state["report_data"][metric]["insight"])

        chart_type = st.selectbox(
            f"Choose chart type for {metric}:", ["Bar", "Line", "Pie", "Table"],
            key=f"chart_select_{metric}",
            index=["Bar", "Line", "Pie", "Table"].index(st.session_state["report_data"][metric]["chart_type"])
        )
        st.session_state["report_data"][metric]["chart_type"] = chart_type

        fig, ax = plt.subplots()
        try:
            render_chart(metric, chart_type, ax)
            plt.tight_layout()
            chart_path = f"chart_{metric.replace(' ', '_')}.png"
            plt.savefig(chart_path)
            st.image(chart_path)
            plt.close(fig)
            st.session_state["report_data"][metric]["chart_path"] = chart_path
        except Exception as e:
            st.error(f"Chart generation failed for {metric}: {e}")

    final_pdf = create_pdf_report_with_charts({
        m: (v["insight"], v.get("chart_path", "")) for m, v in st.session_state["report_data"].items()
    })
    st.session_state["final_pdf"] = final_pdf
    st.download_button("📥 Download Report", open(final_pdf, "rb"), file_name=final_pdf)

if st.session_state["final_pdf"]:
    st.subheader("📬 Email Report")
    email = st.text_input("Enter email address:")
    freq = st.selectbox("Send this report regularly?", ["No", "Daily", "Weekly", "Monthly"])
    if st.button("📤 Send Email"):
        if email:
            sent = send_email_with_pdf(email, "AI Report", "Find your report attached.", st.session_state["final_pdf"])
            if sent:
                st.success("✅ Email sent.")
                if freq != "No":
                    st.info(f"⏰ Scheduled for {freq}. (Scheduler backend required.)")
        else:
            st.warning("Enter a valid email.")
