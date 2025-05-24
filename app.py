import os

import datetime
import psycopg2
import streamlit as st
import pandas as pd

import openai
from dotenv import load_dotenv
import re
from pdfminer.high_level import extract_text as extract_pdf_text
from docx import Document

from io import BytesIO
import imaplib
from zipfile import ZipFile
import email

# Streamlit page config
st.set_page_config(page_title="AI Recruitment", layout="wide")

# Custom CSS for modern, beautiful frontend with consistent blue theme
st.markdown("""
<style>
    /* Main background and font settings */
    body {
        font-family: 'Inter', sans-serif;
        color: #1e293b;
        background-color: #f8fafc;
        margin: 0;
        padding: 0;
    }

    /* Main content area */
    .block-container {
        background-color: #ffffff;
        border-radius: 12px;
        padding: 2rem;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
        margin: 1rem auto;
        max-width: 1280px;
    }

    /* Header */
    .header {
        background: linear-gradient(135deg, #1e3a8a, #2563eb);
        color: white;
        padding: 1.5rem;
        border-radius: 12px;
        text-align: center;
        margin-bottom: 2rem;
        box-shadow: 0 6px 12px rgba(0, 0, 0, 0.15);
    }
    .header h1 {
        margin: 0;
        font-size: 2.25rem;
        font-weight: 700;
        letter-spacing: -0.025em;
    }

    /* Sidebar */
    .sidebar .sidebar-content {
        background-color: #ffffff;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
        margin: 1rem;
    }

    /* All buttons - consistent blue theme */
    .stButton>button, 
    .stDownloadButton>button,
    div[data-testid="stForm"]>div>button,
    button[kind="primary"],
    button[kind="secondary"],
    button[kind="formSubmit"],
    div[data-testid="stForm"] button,
    div[data-testid="stForm"]>div>div>button {
        background: linear-gradient(90deg, #2563eb, #3b82f6) !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.75rem 1.5rem !important;
        font-weight: 500 !important;
        font-size: 1rem !important;
        transition: all 0.2s ease !important;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1) !important;
    }
    
    .stButton>button:hover, 
    .stDownloadButton>button:hover,
    div[data-testid="stForm"]>div>button:hover,
    button[kind="primary"]:hover,
    button[kind="secondary"]:hover,
    button[kind="formSubmit"]:hover,
    div[data-testid="stForm"] button:hover,
    div[data-testid="stForm"]>div>div>button:hover {
        background: linear-gradient(90deg, #1e3a8a, #2563eb) !important;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2) !important;
        transform: translateY(-1px);
    }

    /* File uploader */
    .stFileUploader>div>div {
        border: 2px dashed #bfdbfe !important;
        background-color: #f0f8ff !important;
        border-radius: 12px;
        padding: 2rem 1rem;
        transition: all 0.3s ease;
    }
    .stFileUploader>div>div:hover {
        border-color: #2563eb !important;
        background-color: #e0f2fe !important;
    }

    /* Input fields */
    .stTextInput>div>div>input,
    .stTextArea>div>div>textarea,
    .stSelectbox>div>select,
    .stDateInput>div>div>input {
        border: 1px solid #bfdbfe !important;
        border-radius: 8px !important;
        padding: 0.75rem !important;
        background-color: #f8fafc !important;
    }
    .stTextInput>div>div>input:focus,
    .stTextArea>div>div>textarea:focus,
    .stSelectbox>div>select:focus,
    .stDateInput>div>div>input:focus {
        border-color: #2563eb !important;
        box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1) !important;
    }

    /* Expanders */
    .stExpander {
        background-color: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 12px;
        margin-bottom: 1rem;
        overflow: hidden;
    }
    .stExpander:hover {
        border-color: #2563eb;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
    }
    .stExpander .streamlit-expanderHeader {
        font-weight: 600;
        color: #1e3a8a;
    }
    .stExpanderContent {
        background-color: #f8fafc;
        border-radius: 0 0 12px 12px;
        padding: 1rem;
        max-width: 100%;
    }

    /* Change Password Form styling */
    div[data-testid="stForm"][data-testid="change_password_form"] {
        background-color: #f8fafc !important;
        border: 1px solid #dbeafe !important;
        border-radius: 16px !important;
        padding: 4rem 3rem !important;
        min-width: 800px !important;
        width: 100% !important;
        max-width: 900px !important;
        box-sizing: border-box !important;
        overflow-wrap: break-word !important;
        overflow: auto !important;
        box-shadow: 0 4px 10px rgba(0, 0, 0, 0.05) !important;
        margin: 2rem auto !important;
    }

    /* Form inputs */
    div[data-testid="stForm"][data-testid="change_password_form"] .stTextInput {
        margin-bottom: 3rem !important;
    }
    div[data-testid="stForm"][data-testid="change_password_form"] .stTextInput > div > div > input {
        width: 100% !important;
        box-sizing: border-box !important;
        padding: 2rem 1.5rem !important;
        font-size: 1.3rem !important;
        background-color: #ffffff !important;
        border: 1px solid #e0e7ff !important;
        border-radius: 12px !important;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05) !important;
    }

    /* Form submit button */
    div[data-testid="stForm"][data-testid="change_password_form"] > div > button {
        width: 100% !important;
        margin-top: 3rem !important;
        padding: 2rem 1.5rem !important;
        font-size: 1.5rem !important;
        border-radius: 12px !important;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1) !important;
    }

    /* Responsive design */
    @media (max-width: 768px) {
        div[data-testid="stForm"][data-testid="change_password_form"] {
            min-width: 100% !important;
            max-width: 100% !important;
            padding: 2.5rem 1.5rem !important;
            margin: 1rem auto !important;
        }
        div[data-testid="stForm"][data-testid="change_password_form"] .stTextInput > div > div > input {
            padding: 1.25rem 1rem !important;
            font-size: 1.15rem !important;
        }
        div[data-testid="stForm"][data-testid="change_password_form"] > div > button {
            padding: 1.25rem 1rem !important;
            font-size: 1.3rem !important;
        }
    }
</style>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
""", unsafe_allow_html=True)


# PostgreSQL config
DB_CONFIG = {
    'host': os.environ['PGHOST'],
    'database': os.environ['PGDATABASE'],
    'user': os.environ['PGUSER'],
    'password': os.environ['PGPASSWORD'],
    'port': int(os.environ['PGPORT']),
}



load_dotenv()

def get_connection():
    return psycopg2.connect(**DB_CONFIG)

# Constants

RESUME_FOLDER = 'Resumes'

JD_FOLDER = 'JDs'

DATABASE = os.getenv("DATABASE_URL")



EMAIL = os.getenv("EMAIL")

PASSWORD = os.getenv("PASSWORD")

IMAP_SERVER = "imap.stackmail.com"
IMAP_PORT = 993

# --- IMAP Authentication Function ---
def authenticate_imap():
    imap = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    imap.login(EMAIL, PASSWORD)
    return imap

def search_emails(imap, subject_text="", after_date="", before_date=""):
    from datetime import datetime

    def format_date(d):
        return d.strftime("%d-%b-%Y")

    imap.select("inbox")
    criteria = []

    if subject_text:
        criteria.append(f'SUBJECT "{subject_text}"')
    if after_date:
        criteria.append(f'SINCE {format_date(after_date)}')
    if before_date:
        criteria.append(f'BEFORE {format_date(before_date)}')

    status, messages = imap.search(None, *criteria)
    return messages[0].split() if status == "OK" else []

# Function to generate the Excel file
def export_to_excel(data):
    df = pd.DataFrame(data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Resume Analysis')
    output.seek(0)
    return output

# Function to sanitize file names by removing invalid characters
def sanitize_filename(filename):
    filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
    filename = filename.replace("\r", "_").replace("\n", "_")
    return filename

def download_attachments_from_imap(imap, email_ids, destination_folder):
    downloaded = 0
    for email_id in email_ids:
        status, msg_data = imap.fetch(email_id, "(RFC822)")
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                for part in msg.walk():
                    content_disposition = str(part.get("Content-Disposition"))
                    if "attachment" in content_disposition:
                        filename = part.get_filename()
                        if filename:
                            sanitized_filename = sanitize_filename(filename)
                            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                            base, ext = os.path.splitext(sanitized_filename)
                            unique_filename = f"{base}_{timestamp}{ext}"
                            filepath = os.path.join(destination_folder, unique_filename)
                            with open(filepath, "wb") as f:
                                f.write(part.get_payload(decode=True))
                                print(f"Downloaded {unique_filename} to {filepath}")
                                downloaded += 1
    return downloaded

# --- Resume Extraction ---
EMAIL_REGEX = r'[A-Za-z0-9._%+-]+@(?:[A-Za-z0-9-]+\.)+[A-Za-z]{2,}'
MOBILE_REGEX = r'(?:\+?92[\s-]?|0|92[\s-]?)?(?:\(?\d{3}\)?[\s-]?\d{3}[\s-]?\d{4}|\d{10})\b'
NAME_REGEX = r'\b(?:[A-Z][a-z]+|[A-Z]{2,})(?:\s(?:[A-Z][a-z]+|[A-Z]{2,})){1,3}\b'

def extract_text_from_docx(path):
    doc = Document(path)
    return '\n'.join([para.text for para in doc.paragraphs])

def extract_info_from_text(text):
    email = re.findall(EMAIL_REGEX, text)
    mobile = re.findall(MOBILE_REGEX, text)
    names = re.findall(NAME_REGEX, text)
    return {
        'name': names[0] if names else 'Not found',
        'email': email[0] if email else 'Not found',
        'mobile': mobile[0] if mobile else 'Not found'
    }

def extract_job_title_from_filename(jd_path):
    filename = os.path.basename(jd_path)
    if "application for" in filename.lower():
        return filename.split("for", 1)[-1].replace('.docx', '').replace('.doc', '').replace('.pdf', '').strip()
    return "Not found"

def extract_resume_info(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext == '.pdf':
            text = extract_pdf_text(file_path)
        elif ext == '.docx':
            text = extract_text_from_docx(file_path)
        else:
            raise ValueError("Unsupported file type. Only PDF and DOCX are supported.")
        info = extract_info_from_text(text)
        info['file_name'] = os.path.basename(file_path)
        info['text'] = text
        return info
    except Exception:
        return None

def format_date(d):
    return d.strftime("%d-%b-%Y")

def analyze_resume_with_gpt(resume_info, job_description):
    openai.api_key = os.getenv("OPENAI_API_KEY")
    if not openai.api_key:
        st.warning("OpenAI API key not found in environment variables.")
        return "Score: 0\nRecommendation: Analysis failed due to missing API key\nStrengths: None\nGaps: None"

    resume_text = resume_info.get('text', '')
    if not resume_text:
        resume_text = f"Name: {resume_info.get('name', 'Not found')}\nEmail: {resume_info.get('email', 'Not found')}\nMobile: {resume_info.get('mobile', 'Not found')}"

    prompt = f"""
You are an expert HR recruiter specializing in data science hiring. Your task is to critically evaluate a candidate's resume against a job description and assign a realistic score out of 10.

Job Description:
{job_description}

Candidate Resume:
{resume_text}

Instructions:
You are an expert technical recruiter tasked with evaluating a candidate's fit for a data science role based on their resume and the job description. Follow the steps below to conduct a professional screening.

1. Experience Evaluation:
   - Prioritize full-time experience if the job description explicitly requires it (e.g., “3+ years full-time experience”).
   - If not explicitly stated, include internships, freelance work, academic projects, or part-time roles that demonstrate practical exposure.
   - Assess the relevance, duration, and depth of the candidate’s roles to the JD.

2. Skills Assessment:
   - Match candidate skills (technical tools, programming languages, frameworks, platforms) against those mentioned in the JD.
   - Consider both hands-on experience and conceptual understanding.
   - Highlight any unique or in-demand tools (e.g., ML frameworks, cloud platforms, big data tools).

3. Education & Certifications:
   - Evaluate the candidate’s academic background (degree level, institution, major) in relation to the job’s expectations.
   - Note relevant certifications (e.g., AWS, Azure, Google Cloud, Data Science Specializations) that add value.

4. Scoring Guidelines (Rate from 0–10):
   - 8–10: Excellent fit — candidate meets or exceeds key requirements and is job-ready.
   - 5–7: Moderate fit — good potential but needs minor upskilling or experience depth.
   - 0–4: Poor fit — lacks major qualifications or relevant experience.

5. Output Format (Structured and Concise):
Score: [e.g., 8.5]
Recommendation: [e.g., Strong match for interview shortlist.]
Strengths: [e.g., Robust experience with Python, SQL, and end-to-end ML workflows.]
Gaps: [e.g., Limited cloud deployment and stakeholder communication.]

Be objective and realistic. Focus on job readiness, not just keywords. Avoid inflating scores.
"""

    try:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3,
            max_tokens=500
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        st.warning(f"Error: {e}")
        return "Score: 0\nRecommendation: Analysis failed due to an error\nStrengths: None\nGaps: None"


def init_db():
    conn = get_connection()
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS analysis (
        id SERIAL PRIMARY KEY,
        name TEXT,
        email TEXT,
        mobile TEXT,
        strengths TEXT,
        gaps TEXT,
        recommendation TEXT,
        score REAL,
        status TEXT,
        resume_path TEXT,
        job_title TEXT,
        date_added DATE DEFAULT CURRENT_DATE,
        batch_id TEXT,
        UNIQUE(name, email, job_title, batch_id)
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS admin (
        username TEXT PRIMARY KEY,
        password TEXT
    )''')
    c.execute("SELECT * FROM admin WHERE username = %s", ("admin",))
    if not c.fetchone():
        c.execute("INSERT INTO admin (username, password) VALUES (%s, %s)", ("admin", "123"))
    conn.commit()
    conn.close()

def store_analysis(name, email, mobile, strengths, score, recommendation, gaps, resume_path, job_title, batch_id):
    status = "Shortlisted" if float(score) >= 7 else "Rejected"
    conn = get_connection()
    c = conn.cursor()

    c.execute('''
        INSERT INTO analysis 
        (name, email, mobile, strengths, gaps, recommendation, score, status, resume_path, job_title, date_added, batch_id)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, CURRENT_DATE, %s)
        ON CONFLICT (name, email, job_title, batch_id) DO NOTHING
    ''', (name, email, mobile, strengths, gaps, recommendation, score, status, resume_path, job_title, batch_id))

    conn.commit()
    conn.close()
    return "added"



def is_resume_processed(resume_path, job_title, batch_id):
    conn = get_connection()
    c = conn.cursor()
    c.execute('SELECT COUNT(*) FROM analysis WHERE resume_path = %s AND job_title = %s AND batch_id = %s', 
              (resume_path, job_title, batch_id))
    count = c.fetchone()[0]
    conn.close()
    return count > 0

def load_data():
    try:
        conn = get_connection()
        df = pd.read_sql_query("SELECT * FROM analysis ORDER BY id DESC LIMIT 50", conn)
        conn.close()
        return df
    except Exception as e:
        st.error(f"Failed to load data: {e}")
        return pd.DataFrame()


def normalize_folder_name(text):
    return re.sub(r'\W+', '_', text.strip().lower())

# --- Streamlit UI ---
# --- Streamlit UI ---
init_db()
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "username" not in st.session_state:
    st.session_state.username = None
if "page" not in st.session_state:
    st.session_state.page = "dashboard"

if not st.session_state.logged_in:
    st.markdown('<div class="header"><h1>AI Recruitment</h1></div>', unsafe_allow_html=True)
    st.title("Login")
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        if st.form_submit_button("Login"):
            conn = get_connection()
            c = conn.cursor()
            c.execute("SELECT * FROM admin WHERE username=%s AND password=%s", (username, password))
            result = c.fetchone()
            conn.close()
            if result:
                st.session_state.logged_in = True
                st.session_state.username = username
                st.success("Login successful")
                st.rerun()
            else:
                st.error("Invalid credentials")
    st.stop()

# Sidebar
st.sidebar.title("AI Recruitment")
if st.session_state.page != "change_password":
    page = st.sidebar.radio("Navigation", ["Dashboard", "Process Email", "Quick Analysis"], key="nav_radio")
    st.session_state.page = page.lower().replace(" ", "_")
else:
    st.sidebar.radio("Navigation", ["Dashboard", "Process Email", "Quick Analysis"], key="nav_radio", disabled=True)

if st.sidebar.button("Logout"):
    st.session_state.logged_in = False
    st.session_state.username = None
    st.session_state.page = "dashboard"
    st.rerun()

with st.sidebar.expander("Change Password", expanded=True):
    if st.button("Change Password"):
        st.session_state.page = "change_password"
        st.rerun()

def change_password_page():
    st.markdown("<script>window.scrollTo(0, 0);</script>", unsafe_allow_html=True)
    st.markdown('<div class="header"><h1>Change Password</h1></div>', unsafe_allow_html=True)

    with st.container():
        with st.form("change_password_form", clear_on_submit=True):
            current_password = st.text_input("Current Password", type="password")
            new_password = st.text_input("New Password", type="password")
            confirm_password = st.text_input("Confirm New Password", type="password")
            submit_button = st.form_submit_button("Update Password", use_container_width=True)

            if submit_button:
                conn = get_connection()
                c = conn.cursor()
                if st.session_state.username:
                    c.execute("SELECT * FROM admin WHERE username=%s AND password=%s",
                              (st.session_state.username, current_password))
                    result = c.fetchone()
                    if result:
                        if new_password == confirm_password:
                            c.execute("UPDATE admin SET password=%s WHERE username=%s",
                                      (new_password, st.session_state.username))
                            conn.commit()
                            st.success("Password updated successfully. Please log in again.")
                            st.session_state.logged_in = False
                            st.session_state.username = None
                            st.session_state.page = "dashboard"
                            st.rerun()
                        else:
                            st.error("New passwords do not match.")
                    else:
                        st.error("Current password is incorrect.")
                else:
                    st.error("No user session found.")
                conn.close()

    if st.button("Back"):
        st.session_state.page = "dashboard"
        st.rerun()

# Header
if st.session_state.page != "change_password":
    st.markdown('<div class="header"><h1>AI Recruitment</h1></div>', unsafe_allow_html=True)

if st.session_state.page == "change_password":
    change_password_page()

elif st.session_state.page == "dashboard":
    st.title("Recruitment Dashboard")
    # Initialize session state for dates and filtered results if not set
    if "gmail_start_date" not in st.session_state:
        st.session_state.gmail_start_date = datetime.date.today() - datetime.timedelta(days=30)
    if "gmail_end_date" not in st.session_state:
        st.session_state.gmail_end_date = datetime.date.today()
    if "filtered_df" not in st.session_state:
        st.session_state.filtered_df = None

    with st.form("filter_form"):
        col1, col2, col3 = st.columns([1, 1, 1.5])
        start_date = col1.date_input("Resume Evaluation Start Date", value=st.session_state.gmail_start_date)
        end_date = col2.date_input("Resume Evaluation End Date", value=st.session_state.gmail_end_date)
        subject_filter = col3.text_input("Filter by Job Title", value="")
        batch_start_date = col3.date_input("Hiring Start Date", value=None)
        batch_end_date = col3.date_input("Hiring End Date", value=None)
        status_filter = col1.selectbox("Status", ["All", "Shortlisted"], index=0, key="status_filter")
        show_top_n_input = col3.text_input("Show Top N Scorers (Enter a number, 0 for all)", value="0")
        submit_button = st.form_submit_button("Show Results")

    # Validate show_top_n input
    try:
        show_top_n = int(show_top_n_input) if show_top_n_input.strip() else 0
        if show_top_n < 0:
            st.error("Please enter a non-negative number for Top N Scorers.")
            show_top_n = 0
    except ValueError:
        st.error("Please enter a valid number for Top N Scorers.")
        show_top_n = 0

    if submit_button:
        df = load_data()
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    
        if 'date_added' in df.columns:
            df['date_added'] = pd.to_datetime(df['date_added'], errors='coerce')
            filtered_df = df[
                (df['date_added'].dt.date >= start_date) &
                (df['date_added'].dt.date <= end_date)
            ]
        else:
            filtered_df = df
        
        if subject_filter:
            filtered_df = filtered_df[filtered_df['job_title'].str.contains(subject_filter, case=False, na=False, regex=True)]
        
        if batch_start_date and batch_end_date:
            batch_id = f"{batch_start_date.strftime('%Y-%m-%d')}_to_{batch_end_date.strftime('%Y-%m-%d')}"
            filtered_df = filtered_df[filtered_df['batch_id'].str.contains(batch_id, case=False, na=False, regex=True)]
        
        if status_filter != "All":
            filtered_df = filtered_df[filtered_df['status'] == status_filter]
    
        if show_top_n > 0:
            filtered_df = filtered_df.sort_values('score', ascending=False).head(show_top_n)

        # Store filtered results in session state
        st.session_state.filtered_df = filtered_df

    # Display results from session state if available
    if st.session_state.filtered_df is not None:
        filtered_df = st.session_state.filtered_df

        mcol1, mcol2, mcol3 = st.columns(3)
        with mcol1:
            st.markdown('<div class="metric-card metric-card-total">', unsafe_allow_html=True)
            st.metric("Total Resumes", len(filtered_df))
            st.markdown('</div>', unsafe_allow_html=True)
        with mcol2:
            st.markdown('<div class="metric-card metric-card-shortlisted">', unsafe_allow_html=True)
            st.metric("Shortlisted", len(filtered_df[filtered_df['status'] == "Shortlisted"]))
            st.markdown('</div>', unsafe_allow_html=True)
        with mcol3:
            st.markdown('<div class="metric-card metric-card-rejected">', unsafe_allow_html=True)
            st.metric("Rejected", len(filtered_df[filtered_df['status'] == "Rejected"]))
            st.markdown('</div>', unsafe_allow_html=True)
        
        if not filtered_df.empty:
            for index, row in filtered_df.iterrows():
                with st.expander(f"Report - {row['name']} ({row['job_title']}) - Hiring: {row['batch_id']}"):
                    st.markdown('<div class="expander-content">', unsafe_allow_html=True)
                    col1, col2 = st.columns([1, 3])
                    col1.markdown('<span class="label">Name</span>', unsafe_allow_html=True)
                    col2.markdown(f'<span class="value">{row["name"]}</span>', unsafe_allow_html=True)
                    col1, col2 = st.columns([1, 3])
                    col1.markdown('<span class="label">Email</span>', unsafe_allow_html=True)
                    col2.markdown(f'<span class="value">{row["email"]}</span>', unsafe_allow_html=True)
                    col1, col2 = st.columns([1, 3])
                    col1.markdown('<span class="label">Mobile</span>', unsafe_allow_html=True)
                    col2.markdown(f'<span class="value">{row["mobile"]}</span>', unsafe_allow_html=True)
                    col1, col2 = st.columns([1, 3])
                    col1.markdown('<span class="label">Score</span>', unsafe_allow_html=True)
                    col2.markdown(f'<span class="value">{row["score"]}</span>', unsafe_allow_html=True)
                    col1, col2 = st.columns([1, 3])
                    col1.markdown('<span class="label">Recommendation</span>', unsafe_allow_html=True)
                    col2.markdown(f'<span class="value">{row["recommendation"]}</span>', unsafe_allow_html=True)
                    col1, col2 = st.columns([1, 3])
                    col1.markdown('<span class="label">Gaps</span>', unsafe_allow_html=True)
                    col2.markdown(f'<span class="value">{row["gaps"]}</span>', unsafe_allow_html=True)
                    col1, col2 = st.columns([1, 3])
                    col1.markdown('<span class="label">Strengths</span>', unsafe_allow_html=True)
                    col2.markdown(f'<span class="value">{row.get("strengths", "Not Available")}</span>', unsafe_allow_html=True)
                    col1, col2 = st.columns([1, 3])
                    col1.markdown('<span class="label">Status</span>', unsafe_allow_html=True)
                    col2.markdown(f'<span class="value">{row["status"]}</span>', unsafe_allow_html=True)
                    col1, col2 = st.columns([1, 3])
                    col1.markdown('<span class="label">Job Title</span>', unsafe_allow_html=True)
                    col2.markdown(f'<span class="value">{row["job_title"]}</span>', unsafe_allow_html=True)
                    col1, col2 = st.columns([1, 3])
                    col1.markdown('<span class="label">Batch ID</span>', unsafe_allow_html=True)
                    col2.markdown(f'<span class="value">{row["batch_id"]}</span>', unsafe_allow_html=True)
                    resume_path = row.get('resume_path', None)
                    if resume_path and os.path.exists(resume_path):
                        with open(resume_path, "rb") as file:
                            st.download_button(
                                label="📄 Download Resume",
                                data=file,
                                file_name=os.path.basename(resume_path),
                                mime="application/octet-stream",
                                key=f"download_resume_{index}"
                            )
                    else:
                        st.info("ℹ️ Resume file not found, but extracted details are available in the report.")
                    st.markdown('</div>', unsafe_allow_html=True)

        # Add a single export button at the end of all filtered results
        if not filtered_df.empty:
            # Create an Excel export of the entire filtered DataFrame
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                filtered_df.to_excel(writer, index=False, sheet_name='Filtered Resumes')
            excel_data = output.getvalue()

            # Display the download button for the whole filtered dataset
            st.download_button(
                label="📊 Export All to Excel",
                data=excel_data,
                file_name="filtered_resumes.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("No results found matching the filters.")


elif st.session_state.page == "process_email":
    st.title("Process Email Resumes")

    job_keyword = st.text_input("Enter Job Keyword (e.g., Data Scientist)").strip()
    col1, col2 = st.columns(2)
    start_date = col1.date_input("Start Date", datetime.date.today() - datetime.timedelta(days=7))
    end_date = col2.date_input("End Date", datetime.date.today())

    if st.button("Fetch Resumes"):
        if not job_keyword:
            st.error("Please enter a job keyword.")
        else:
            folder_name = normalize_folder_name(job_keyword)
            date_suffix = f"{start_date.strftime('%Y-%m-%d')}_to_{end_date.strftime('%Y-%m-%d')}"
            resume_subfolder = os.path.join(RESUME_FOLDER, folder_name, date_suffix)
            os.makedirs(resume_subfolder, exist_ok=True)

            try:
                imap = authenticate_imap()
                email_ids = search_emails(imap, subject_text=job_keyword, after_date=start_date, before_date=end_date + datetime.timedelta(days=1))
                if not email_ids:
                    st.warning("No emails matched the job keyword in subject line.")
                else:
                    downloaded = download_attachments_from_imap(imap, email_ids, resume_subfolder)
                    st.success(f"Downloaded {downloaded} resumes to {resume_subfolder}.")
                imap.logout()
            except imaplib.IMAP4.error as e:
                st.error(f"IMAP Login failed: {str(e)}")

    st.subheader("Job Description Input")
    jd_input_method = st.radio("Provide Job Description by:", ("Upload File(s)", "Paste Text"))
    jd_text_input = ""
    uploaded_files = []

    if jd_input_method == "Upload File(s)":
        uploaded_files = st.file_uploader("Upload JD file(s)", type=["txt", "docx", "pdf"], accept_multiple_files=True)
        if uploaded_files:
            for uploaded_file in uploaded_files:
                save_path = os.path.join(JD_FOLDER, uploaded_file.name)
                with open(save_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
            st.success(f"Uploaded {len(uploaded_files)} JD file(s) to {JD_FOLDER}")
    else:
        jd_text_input = st.text_area("Paste the Job Description text here")

    if st.button("Process Resumes"):
        with st.spinner("Processing resumes..."):
            folder_name = normalize_folder_name(job_keyword)
            date_suffix = f"{start_date.strftime('%Y-%m-%d')}_to_{end_date.strftime('%Y-%m-%d')}"
            batch_id = date_suffix
            resume_subfolder = os.path.join(RESUME_FOLDER, folder_name, date_suffix)

            if not os.path.exists(resume_subfolder):
                st.error(f"No resume folder found for keyword '{job_keyword}' at {resume_subfolder}")
            else:
                if jd_input_method == "Upload File(s)":
                    jd_files = [f for f in os.listdir(JD_FOLDER) if os.path.isfile(os.path.join(JD_FOLDER, f))]
                    if not jd_files:
                        st.error("No JD files uploaded.")
                        st.stop()
                    jd_path = os.path.join(JD_FOLDER, jd_files[0])
                    ext = os.path.splitext(jd_path)[1].lower()
                    if ext == '.txt':
                        job_description = open(jd_path, 'r', encoding='utf-8').read()
                    elif ext == '.docx':
                        job_description = extract_text_from_docx(jd_path)
                    elif ext == '.pdf':
                        job_description = extract_pdf_text(jd_path)
                    else:
                        st.error("Unsupported JD file format.")
                        st.stop()
                else:
                    job_description = jd_text_input.strip()

                if not job_description:
                    st.error("Job description content is empty.")
                    st.stop()

                total_processed = 0
                total_failed = 0
                total_duplicates = 0
                job_title = job_keyword.title()

                for filename in os.listdir(resume_subfolder):
                    resume_path = os.path.join(resume_subfolder, filename)
                    if is_resume_processed(resume_path, job_title, batch_id):
                        continue

                    resume_info = extract_resume_info(resume_path)
                    if not resume_info or resume_info['name'] == 'Not found':
                        total_failed += 1
                        continue

                    result = analyze_resume_with_gpt(resume_info, job_description)
                    if not result:
                        total_failed += 1
                        continue

                    score = 0
                    strengths = ""
                    recommendation = ""
                    gaps = ""

                    for line in result.splitlines():
                        if "score" in line.lower():
                            try:
                                match = re.search(r'score.*?:\s*(\d+\.?\d*)', line, re.IGNORECASE)
                                if match:
                                    score = float(match.group(1))
                            except:
                                pass
                        elif "strengths" in line.lower():
                            strengths = line.split(":", 1)[-1].strip()
                        elif "recommendation" in line.lower():
                            recommendation = line.split(":", 1)[-1].strip()
                        elif "gap" in line.lower():
                            gaps = line.split(":", 1)[-1].strip()

                    name = resume_info.get('name', 'Not found')
                    email = resume_info.get('email', 'Not found')
                    mobile = resume_info.get('mobile', 'Not found')

                    store_result = store_analysis(
                        name, email, mobile,
                        strengths, score, recommendation, gaps,
                        resume_path, job_title, batch_id
                    )

                    if store_result == "added":
                        total_processed += 1
                    elif store_result == "duplicate":
                        total_duplicates += 1

                st.success(
                    f"Job: {job_title} → Processed: {total_processed}, Duplicates Skipped: {total_duplicates}, Failed: {total_failed}"
                )


elif st.session_state.page == "quick_analysis":
    st.title("Quick Resume Analysis")
    # Initialize session state for results if not set
    if "quick_analysis_results" not in st.session_state:
        st.session_state.quick_analysis_results = []

    jd_input_method = st.radio("How would you like to provide the Job Description?", ("Upload File", "Paste Text"))
    uploaded_jd = None
    jd_text_input = ""

    if jd_input_method == "Upload File":
        uploaded_jd = st.file_uploader("Upload Job Description", type=["pdf", "doc", "docx"])
    else:
        jd_text_input = st.text_area("Paste Job Description text here")

    uploaded_resumes = st.file_uploader("Upload Resumes", type=["pdf", "doc", "docx"], accept_multiple_files=True)

    if st.button("Process Resumes"):
        with st.spinner("Processing resumes..."):
            try:
                os.makedirs(JD_FOLDER, exist_ok=True)
                os.makedirs(RESUME_FOLDER, exist_ok=True)

                if jd_input_method == "Paste Text" and jd_text_input.strip():
                    jd_text = jd_text_input.strip()
                    job_title = "Job Description"
                elif jd_input_method == "Upload File" and uploaded_jd:
                    jd_path = os.path.join(JD_FOLDER, uploaded_jd.name)
                    with open(jd_path, "wb") as f:
                        f.write(uploaded_jd.read())
                    if jd_path.endswith(('.docx', '.doc')):
                        jd_text = extract_text_from_docx(jd_path)
                    elif jd_path.endswith('.pdf'):
                        jd_text = extract_pdf_text(jd_path)
                    else:
                        raise ValueError("Unsupported JD format.")
                    job_title = extract_job_title_from_filename(jd_path)
                else:
                    st.error("Please provide a Job Description file or paste the text.")
                    st.stop()

                if not uploaded_resumes:
                    st.error("Please upload at least one Resume.")
                    st.stop()

                results = []
                for uploaded_resume in uploaded_resumes:
                    file_name = uploaded_resume.name
                    file_bytes = uploaded_resume.read()
                    # Upload to Supabase
                    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    base, ext = os.path.splitext(file_name)
                    unique_filename = f"{base}_{timestamp}{ext}"
                    supabase.storage.from_('resumes').upload(f"quick_analysis/{unique_filename}", file_bytes)
                    resume_url = supabase.storage.from_('resumes').get_public_url(f"quick_analysis/{unique_filename}")

                    resume_info = extract_resume_info(resume_url)
                    if not resume_info:
                        continue

                    result = analyze_resume_with_gpt(resume_info, jd_text)
                    if not result:
                        continue

                    score = 0
                    recommendation, gaps, strengths = "", "", ""
                    for line in result.splitlines():
                        if "score" in line.lower():
                            try:
                                match = re.search(r'score.*?:\s*(\d+\.?\d*)', line, re.IGNORECASE)
                                score = float(match.group(1)) if match else 0
                            except:
                                score = 0
                        elif "recommendation" in line.lower():
                            recommendation = line.split(":", 1)[-1].strip()
                        elif "gap" in line.lower():
                            gaps = line.split(":", 1)[-1].strip()
                        elif "strength" in line.lower():
                            strengths = line.split(":", 1)[-1].strip()

                    status = "Shortlisted" if score >= 7 else "Rejected"

                    results.append({
                        "name": resume_info.get("name", "Not found"),
                        "email": resume_info.get("email", "Not found"),
                        "mobile": resume_info.get("mobile", "Not found"),
                        "score": score,
                        "recommendation": recommendation,
                        "gaps": gaps,
                        "strengths": strengths,
                        "resume_path": resume_url,
                        "job_title": job_title,
                        "status": status,
                    })

                # Store results in session state
                st.session_state.quick_analysis_results = results
                st.success("Analysis complete! See results below.")

            except Exception as e:
                st.error(f"Error processing resumes: {e}")

    # Display results from session state if available
    if st.session_state.quick_analysis_results:
        st.subheader("Resume Analysis Results")
        for index, row in enumerate(st.session_state.quick_analysis_results):
            with st.expander(f"Report - {row['name']} ({row['job_title']})"):
                col1, col2 = st.columns([1, 3])
                col1.markdown("**Name**")
                col2.write(row["name"])
                col1, col2 = st.columns([1, 3])
                col1.markdown("**Email**")
                col2.write(row["email"])
                col1, col2 = st.columns([1, 3])
                col1.markdown("**Mobile**")
                col2.write(row["mobile"])
                col1, col2 = st.columns([1, 3])
                col1.markdown("**Score**")
                col2.markdown(f'<span class="value">{row["score"]}</span>', unsafe_allow_html=True)
                col1, col2 = st.columns([1, 3])
                col1.markdown("**Recommendation**")
                col2.write(row["recommendation"])
                col1, col2 = st.columns([1, 3])
                col1.markdown("**Gaps**")
                col2.write(row["gaps"])
                col1, col2 = st.columns([1, 3])
                col1.markdown("**Strengths**")
                col2.write(row["strengths"])
                col1, col2 = st.columns([1, 3])
                col1.markdown("**Status**")
                col2.write(row["status"])
                col1, col2 = st.columns([1, 3])
                col1.markdown("**Job Title**")
                col2.write(row["job_title"])


                df = pd.DataFrame([row])
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Resume Analysis')
                excel_data = output.getvalue()
                col2.download_button(
                    label="📊 Export to Excel",
                    data=excel_data,
                    file_name=f"{row['name']}_resume_analysis.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"export_excel_quick_{index}"
                )


                # Zip all resumes and offer as a single download
                zip_buffer = BytesIO()
                with ZipFile(zip_buffer, "w") as zipf:
                    for row in st.session_state.quick_analysis_results:
                        resume_path = row.get("resume_path")
                        if resume_path and os.path.exists(resume_path):
                            zipf.write(resume_path, arcname=os.path.basename(resume_path))
                zip_buffer.seek(0)

                st.download_button(
                    label="📦 Download All Resumes (ZIP)",
                    data=zip_buffer,
                    file_name=f"{row['job_title'].replace(' ', '_')}_resumes.zip",
                    mime="application/zip",
                    key="download_all_zip"
                )
    else:
        st.info("No resumes analyzed yet.")
