import pandas as pd
import streamlit as st
from dotenv import load_dotenv
import os
from langchain.chat_models import ChatOpenAI
from langchain_experimental.agents import create_pandas_dataframe_agent
from gtts import gTTS
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.audio import MIMEAudio
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime
import time
from fpdf import FPDF

# Load environment variables
load_dotenv()
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

@st.cache_data
def load_data():
    df = pd.read_excel("Dummy_Course_Data.xlsx", engine="openpyxl")
    return df

df = load_data()
df["Course Completion Date"] = pd.to_datetime(df["Course Completion Date"], errors="coerce")
df["Registration Date"] = pd.to_datetime(df["Registration Date"], errors="coerce")

insight_options = {
    "1. Employees with 2 or more completions": "Number of employees who have completed 2 or more courses.",
    "2. Employees with 1 completion": "Number of employees who have completed exactly 1 course.",
    "3. Top 3 organizations by completion": "Top 3 organization units with the highest number of completions.",
    "4. Top 3 orgs in each country by completion": "Top 3 organization units by completions in each country.",
    "5. Bottom 3 organizations by completion": "Bottom 3 organization units with the least completions.",
    "6. Bottom 3 orgs where registration is high but completion is low": "Bottom 3 orgs where registration is high but completion is low.",
    "7. Top 3 platforms by registration and completion": "Top 3 platforms with the highest registrations and completions.",
    "8. Bottom 3 platforms where registration is high but completion is low": "Bottom 3 platforms where registration is high but completion is low.",
    "9. Top 3 employee roles by completions": "Top 3 employee roles based on number of completions.",
    "10. Split of completion by course level": "Split of completions across different course levels.",
    "11. Course level with high registration but low completion": "Which course levels have high registration but low completion?"
}

def run_agent_on_prompt(prompt_text, model="gpt-4"):
    llm = ChatOpenAI(openai_api_key=OPENAI_API_KEY, model_name=model, temperature=0)
    agent = create_pandas_dataframe_agent(
        llm, df, verbose=False, handle_parsing_errors=True, allow_dangerous_code=True
    )
    return agent.run(prompt_text)

def text_to_speech(text, filename="summary.mp3"):
    tts = gTTS(text)
    tts.save(filename)
    return filename

def generate_pdf_from_text(text, filename="auto_insights.pdf"):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font("Arial", size=12)
    for line in text.split("\n"):
        pdf.multi_cell(0, 10, line)
    pdf.output(filename)
    return filename

def send_email_with_attachments(subject, body, audio_path=None, pdf_path=None):
    try:
        sender = os.getenv("EMAIL_USERNAME")
        password = os.getenv("EMAIL_PASSWORD")
        receivers = os.getenv("EMAIL_RECEIVERS").split(",")
        host = os.getenv("EMAIL_HOST")
        port = int(os.getenv("EMAIL_PORT"))

        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = ", ".join(receivers)
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain"))

        if audio_path:
            with open(audio_path, "rb") as f:
                part = MIMEAudio(f.read(), _subtype="mp3")
                part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(audio_path)}")
                msg.attach(part)

        if pdf_path:
            with open(pdf_path, "rb") as f:
                part = MIMEApplication(f.read(), _subtype="pdf")
                part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(pdf_path)}")
                msg.attach(part)

        with smtplib.SMTP(host, port) as server:
            server.starttls()
            server.login(sender, password)
            server.sendmail(sender, receivers, msg.as_string())

        return True
    except Exception as e:
        st.error(f"Email Error: {e}")
        return False

def daily_auto_email():
    try:
        full_summary = ""
        for label, prompt_text in insight_options.items():
            prompt = f"Analyze the dataset and return this insight:\n\n{prompt_text}"
            try:
                result = run_agent_on_prompt(prompt, model="gpt-3.5-turbo")
            except Exception as e:
                result = f"[⚠️ Skipped due to error: {e}]"
            full_summary += f"\n\n### {label}\n{result}\n"

        with open("auto_insights.txt", "w") as f:
            f.write(full_summary)

        podcast_prompt = f"Summarize the following into a podcast-style audio summary:\n\n{full_summary}"
        podcast_text = run_agent_on_prompt(podcast_prompt, model="gpt-4")
        podcast_path = text_to_speech(podcast_text, "auto_podcast.mp3")
        pdf_path = generate_pdf_from_text(full_summary)

        subject = f"📩 Daily Course Insights – {datetime.now().strftime('%Y-%m-%d')}"
        body = "Attached are your latest daily insights and podcast summary."
        success = send_email_with_attachments(subject, body, audio_path=podcast_path, pdf_path=pdf_path)
        if success:
            st.success("✅ Email sent successfully.")
        else:
            st.error("❌ Email failed.")
    except Exception as e:
        st.error(f"❌ Auto email error: {e}")

# ---------------- UI ----------------
st.title("📘 Sheet 1 – AI-Powered Insight Dashboard")

# Insight generation
st.subheader("📊 Select and Generate a Specific Insight")
selected_insight = st.selectbox("Choose Insight", list(insight_options.keys()))
if st.button("🔍 Generate Insight"):
    st.info("Generating insight with GPT-4...")
    try:
        prompt = f"Analyze the dataset and return this insight:\n\n{insight_options[selected_insight]}"
        insight_result = run_agent_on_prompt(prompt, model="gpt-4")
        st.text_area("📄 Insight Result", insight_result, height=300)
        st.download_button("📥 Download Insight", insight_result, file_name="Insight.txt")
        st.session_state["last_insight"] = insight_result
    except Exception as e:
        st.error(f"Error: {e}")

# Podcast for single insight
if "last_insight" in st.session_state:
    st.subheader("🎙️ Create Podcast for This Insight")
    if st.button("🎧 Generate Podcast"):
        try:
            summary_prompt = f"Convert this into a 60-second podcast summary:\n\n{st.session_state['last_insight']}"
            podcast_text = run_agent_on_prompt(summary_prompt, model="gpt-4")
            audio_file = text_to_speech(podcast_text, "podcast_individual.mp3")
            st.audio(open(audio_file, "rb").read(), format="audio/mp3")
            st.success("🎧 Podcast ready.")
        except Exception as e:
            st.error(f"Podcast error: {e}")

# Generate all insights
st.markdown("---")
st.subheader("⚡ Generate All Insights")
if st.button("🧠 Generate All"):
    full_summary = ""
    try:
        for label, text in insight_options.items():
            try:
                prompt = f"Analyze the dataset and return this insight:\n\n{text}"
                result = run_agent_on_prompt(prompt, model="gpt-3.5-turbo")
            except Exception as e:
                result = f"[⚠️ Skipped due to error: {e}]"
            full_summary += f"\n\n### {label}\n{result}\n"

        st.text_area("📋 All Insights", full_summary, height=500)
        st.download_button("📥 Download All", full_summary, file_name="All_Insights.txt")
        st.session_state["all_insights"] = full_summary
    except Exception as e:
        st.error(f"Error: {e}")

# Podcast for full summary
if "all_insights" in st.session_state:
    st.subheader("🎧 Create Podcast for All Insights")
    if st.button("🎙️ Generate Full Podcast"):
        try:
            podcast_prompt = f"Summarize all these insights into a 2-minute podcast:\n\n{st.session_state['all_insights']}"
            podcast_text = run_agent_on_prompt(podcast_prompt, model="gpt-4")
            podcast_path = text_to_speech(podcast_text, "full_podcast.mp3")
            st.audio(open(podcast_path, "rb").read(), format="audio/mp3")
            st.success("✅ Podcast ready.")
        except Exception as e:
            st.error(f"Podcast error: {e}")

# Smart Q&A
st.markdown("---")
st.subheader("🤖 Ask Your Own Question About the Data")
user_q = st.text_input("Type your question:")

if st.button("Ask AI"):
    if user_q:
        st.info("Using GPT-4...")
        try:
            summary_prompt = f"""
You are an AI analyst. Based on the dataset, answer the following question in 5 lines maximum:

{user_q}
"""
            try:
                result = run_agent_on_prompt(summary_prompt, model="gpt-4")
            except Exception as e:
                if "Rate limit" in str(e):
                    st.warning("Rate limit hit. Waiting 20 seconds...")
                    time.sleep(20)
                    result = run_agent_on_prompt(summary_prompt, model="gpt-4")
                else:
                    raise e

            st.text_area("🧠 Insight Summary:", result, height=200)
            st.download_button("📥 Download Insight", result, file_name="QnA_Insight.txt")

        except Exception as e:
            st.error(f"Smart Q&A Error: {e}")
    else:
        st.warning("Please enter a question.")

# Auto Email Now
st.markdown("---")
st.subheader("📩 Auto Email Insights + Podcast")
if st.button("📬 Send Email Now"):
    daily_auto_email()
