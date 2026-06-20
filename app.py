import os
import smtplib
import json
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import streamlit as st

EMAIL_LOG_FILE = "email_logs.json"
ACTION_LOG_FILE = "action_logs.json"

def load_logs():
    email_logs = []
    action_logs = []
    if os.path.exists(EMAIL_LOG_FILE):
        try:
            with open(EMAIL_LOG_FILE, "r") as f:
                email_logs = json.load(f)
        except Exception:
            pass
    if os.path.exists(ACTION_LOG_FILE):
        try:
            with open(ACTION_LOG_FILE, "r") as f:
                action_logs = json.load(f)
        except Exception:
            pass
    return email_logs, action_logs

def save_logs(email_logs, action_logs):
    try:
        with open(EMAIL_LOG_FILE, "w") as f:
            json.dump(email_logs, f, indent=4)
        with open(ACTION_LOG_FILE, "w") as f:
            json.dump(action_logs, f, indent=4)
    except Exception:
        pass

# Set Streamlit Page Configuration
st.set_page_config(
    page_title="Attendance Tracker",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded"
)

# Elegant Custom CSS
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Outfit', sans-serif;
    }
    
    .main-title {
        font-size: 2.5rem;
        font-weight: 700;
        background: linear-gradient(135deg, #3B82F6 0%, #10B981 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.5rem;
    }
    
    .sub-title {
        font-size: 1.1rem;
        color: #6B7280;
        margin-bottom: 2rem;
    }
    
    .metric-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.2rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.05);
    }
    
    .metric-label {
        font-size: 0.85rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #64748b;
        margin-bottom: 0.25rem;
    }
    
    .metric-value {
        font-size: 2rem;
        font-weight: 700;
        color: #0f172a;
    }
    
    .log-box {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        border-left: 4px solid #3B82F6;
        border-radius: 4px;
        padding: 0.75rem 1rem;
        margin-bottom: 0.5rem;
        font-family: monospace;
        font-size: 0.9rem;
        color: #0f172a;
    }
    
    .email-preview {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 1rem;
        margin-top: 0.5rem;
        margin-bottom: 1rem;
        box-shadow: inset 0 2px 4px 0 rgba(0, 0, 0, 0.02);
        color: #0f172a;
    }
    
    .badge {
        display: inline-block;
        padding: 0.25em 0.6em;
        font-size: 75%;
        font-weight: 700;
        line-height: 1;
        text-align: center;
        white-space: nowrap;
        vertical-align: baseline;
        border-radius: 0.25rem;
        margin-right: 0.2rem;
    }
    
    .badge-warning {
        background-color: #F59E0B;
        color: #1E1B4B;
    }
    
    .badge-critical {
        background-color: #EF4444;
        color: #FFFFFF;
    }
    
    .badge-normal {
        background-color: #10B981;
        color: #FFFFFF;
    }
</style>
""", unsafe_allow_html=True)

# Database Management
FILE_PATH = "attendance.xlsx"

def load_data():
    """Loads student database from Excel, initializing if not present."""
    if not os.path.exists(FILE_PATH):
        # Fallback initializer in case it wasn't run
        data = {
            "Roll Number": [1, 2, 3, 4, 5],
            "Student Email": [
                "student1@example.com",
                "student2@example.com",
                "student3@example.com",
                "student4@example.com",
                "student5@example.com"
            ],
            "CI": [0, 1, 2, 0, 1],
            "Python": [0, 1, 2, 1, 2],
            "DM": [0, 1, 2, 2, 0]
        }
        df = pd.DataFrame(data)
        df.to_excel(FILE_PATH, index=False, sheet_name="Sheet1")
    return pd.read_excel(FILE_PATH)

def save_data(df):
    """Saves student dataframe back to Excel."""
    df.to_excel(FILE_PATH, index=False)

# Session State Initialization
if "email_logs" not in st.session_state or "action_logs" not in st.session_state:
    saved_emails, saved_actions = load_logs()
    st.session_state.email_logs = saved_emails
    st.session_state.action_logs = saved_actions

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
    st.session_state.logged_in_user = None

# Load current attendance records
df_attendance = load_data().dropna(subset=["Roll Number"])
df_attendance["Roll Number"] = df_attendance["Roll Number"].astype(int)

# Subjects mapping
SUBJECTS = {
    "Computer Intelligence (CI)": {"col": "CI", "code": 1},
    "Python": {"col": "Python", "code": 2},
    "Data Mining (DM)": {"col": "DM", "code": 3}
}

# Sidebar Layout - Configuration
st.sidebar.markdown("### Tracker Configuration")

# Staff Authentication Gate
STAFF_CREDENTIALS_FILE = "staff_credentials.json"

def load_staff_credentials():
    """Loads staff username/password pairs from file, creating a default admin account if missing."""
    if os.path.exists(STAFF_CREDENTIALS_FILE):
        try:
            with open(STAFF_CREDENTIALS_FILE, "r") as f:
                return json.load(f)
        except Exception:
            pass
    default_creds = {"admin": "admin123"}
    with open(STAFF_CREDENTIALS_FILE, "w") as f:
        json.dump(default_creds, f, indent=4)
    return default_creds

def save_staff_credentials(creds):
    with open(STAFF_CREDENTIALS_FILE, "w") as f:
        json.dump(creds, f, indent=4)

staff_credentials = load_staff_credentials()

st.sidebar.markdown("---")
st.sidebar.markdown("### Staff Login")

if not st.session_state.authenticated:
    auth_tab = st.sidebar.radio("Access", ["Login", "Sign Up"], horizontal=True, label_visibility="collapsed")

    if auth_tab == "Login":
        with st.sidebar.form("login_form"):
            login_username = st.text_input("Username", placeholder="Enter staff username")
            login_password = st.text_input("Password", type="password", placeholder="Enter password")
            login_submit = st.form_submit_button("Login", use_container_width=True)

            if login_submit:
                if login_username in staff_credentials and staff_credentials[login_username] == login_password:
                    st.session_state.authenticated = True
                    st.session_state.logged_in_user = login_username
                    st.rerun()
                else:
                    st.error("Invalid username or password.")
    else:
        with st.sidebar.form("signup_form"):
            new_username = st.text_input("Choose Username", placeholder="e.g. ci_staff")
            new_password = st.text_input("Choose Password", type="password")
            confirm_password = st.text_input("Confirm Password", type="password")
            signup_submit = st.form_submit_button("Create Account", use_container_width=True)

            if signup_submit:
                if not new_username or not new_password:
                    st.error("Username and password cannot be empty.")
                elif new_username in staff_credentials:
                    st.error("This username already exists. Please log in instead.")
                elif new_password != confirm_password:
                    st.error("Passwords do not match.")
                else:
                    staff_credentials[new_username] = new_password
                    save_staff_credentials(staff_credentials)
                    st.success("Account created successfully. You can now log in from the Login tab.")
else:
    st.sidebar.success(f"Logged in as: {st.session_state.logged_in_user}")
    if st.sidebar.button("Logout"):
        st.session_state.authenticated = False
        st.session_state.logged_in_user = None
        st.rerun()

is_authenticated = st.session_state.authenticated

# Threshold configuration
warning_threshold = st.sidebar.number_input("Warning Leave Threshold", min_value=1, max_value=5, value=2, step=1,
                                            help="Sends student a warning email when they reach this number of leaves.")
critical_threshold = st.sidebar.number_input("Critical Leave Threshold", min_value=2, max_value=10, value=3, step=1,
                                             help="Sends a lack of attendance alert to both student and staff when leaves exceed warning threshold (>= critical).")

# Validate threshold ordering so the warning/critical logic stays consistent
if warning_threshold >= critical_threshold:
    st.sidebar.error("Warning threshold must be lower than Critical threshold. Please adjust the values above.")

st.sidebar.markdown("---")
st.sidebar.markdown("### Email Credentials")
simulation_mode = st.sidebar.toggle("Simulation Mode", value=True,
                                    help="When active, emails are simulated and displayed on the dashboard instead of actually being sent.")

if not simulation_mode:
    sender_email = st.sidebar.text_input("Sender Gmail Address", value="", placeholder="your-email@gmail.com")
    sender_password = st.sidebar.text_input("App Password", value="", type="password", placeholder="xxxx xxxx xxxx xxxx",
                                            help="Use a Gmail App Password, NOT your regular password. Requires 2FA enabled on Google account.")
    smtp_server = st.sidebar.text_input("SMTP Server", value="smtp.gmail.com")
    smtp_port = st.sidebar.number_input("SMTP Port", value=587)
else:
    st.sidebar.info("Simulation Mode is active. No real credentials needed. Emails will display in the Activity Logs below.")
    sender_email = "tracker-noreply@school.edu"
    sender_password = ""
    smtp_server = ""
    smtp_port = 0

st.sidebar.markdown("---")
st.sidebar.markdown("### Staff Emails")
ci_staff = st.sidebar.text_input("CI Subject Staff Email", value="ci_staff@school.edu")
python_staff = st.sidebar.text_input("Python Subject Staff Email", value="python_staff@school.edu")
dm_staff = st.sidebar.text_input("DM Subject Staff Email", value="dm_staff@school.edu")

staff_emails = {
    "CI": ci_staff,
    "Python": python_staff,
    "DM": dm_staff
}

st.sidebar.markdown("---")
with st.sidebar.expander("About This Project", expanded=False):
    st.markdown("""
    **Simulation Mode**
    This application is configured in **Simulation Mode** by default. In this mode, email alerts are logged and displayed directly on the screen (in the Email Notifications Log) instead of being sent. This allows testing all functionalities immediately without needing mail server credentials.
    
    **Data Persistence**
    All attendance records are loaded from and saved to **`attendance.xlsx`** (Excel format) in the project root directory. Because the app reads and writes directly to this file, your changes **persist across sessions** (even if you close and restart the server).
    
    **Tech Stack**
    * **Streamlit**: Web frontend framework
    * **Pandas**: Data manipulation and query engine
    * **openpyxl**: Excel spreadsheet file handler
    * **smtplib & email**: Mail transmission protocols
    """)

# Email Notification Engine
def log_email(to_email, subject, body, status="Simulated", details=""):
    """Adds email status log to session state."""
    log_entry = {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "to": to_email,
        "subject": subject,
        "body": body,
        "status": status,
        "details": details
    }
    st.session_state.email_logs.insert(0, log_entry)
    save_logs(st.session_state.email_logs, st.session_state.action_logs)

def send_notification(to_email, subject, body):
    """Sends email notifications (live or simulated)."""
    if simulation_mode:
        log_email(to_email, subject, body, "Simulated", "Simulated transmission successful.")
        return True, "Simulated email successfully logged."
    
    if not sender_email or not sender_password:
        log_email(to_email, subject, body, "Failed", "Sender email or password is empty.")
        return False, "Credentials missing."
        
    try:
        s = smtplib.SMTP(smtp_server, int(smtp_port), timeout=10)
        s.starttls()
        s.login(sender_email, sender_password)
        
        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = to_email
        message['Subject'] = subject
        message.attach(MIMEText(body, 'plain'))
        
        content = message.as_string()
        s.sendmail(sender_email, to_email, content)
        s.quit()
        
        log_email(to_email, subject, body, "Sent", "Delivered via SMTP.")
        return True, "Email sent successfully."
    except Exception as e:
        error_msg = str(e)
        log_email(to_email, subject, body, "Failed", error_msg)
        return False, error_msg

# Main Application Interface

# Header Block
st.markdown("""
<div style="background-color: rgba(255, 255, 255, 0.03); border: 1px solid rgba(255, 255, 255, 0.1); border-radius: 8px; padding: 15px; margin-bottom: 20px;">
    <div style="display: flex; flex-wrap: wrap; justify-content: space-between; font-size: 0.95rem;">
        <div style="flex: 1; min-width: 250px; padding: 5px;">
            <strong>Student Name:</strong> Rajesh Yadav<br>
            <strong>Roll Number:</strong> 2300320130190
        </div>
        <div style="flex: 1; min-width: 250px; padding: 5px;">
            <strong>Course/Department:</strong> B.Tech - Information Technology<br>
            <strong>College Name:</strong> ABES Engineering College
        </div>
        <div style="flex: 1; min-width: 150px; padding: 5px; text-align: right;">
            <strong>Date:</strong> June 20, 2026
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown("<div class='main-title'>Attendance Tracker</div>", unsafe_allow_html=True)
st.markdown("<div class='sub-title'>Log student absences, monitor attendance levels, and automate notification alerts.</div>", unsafe_allow_html=True)

# Section 1: Dashboard Metrics
col1, col2, col3, col4 = st.columns(4)

total_students = len(df_attendance)

# Count warning students (leaves == warning_threshold) and critical students (leaves >= critical_threshold)
# across all subjects
warning_count = 0
critical_count = 0

for subj in ["CI", "Python", "DM"]:
    warning_count += sum(df_attendance[subj] == warning_threshold)
    critical_count += sum(df_attendance[subj] >= critical_threshold)

with col1:
    st.markdown(f"""
    <div class='metric-card'>
        <div class='metric-label'>Total Students</div>
        <div class='metric-value'>{total_students}</div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown(f"""
    <div class='metric-card'>
        <div class='metric-label'>Warnings (Exactly {warning_threshold} Leaves)</div>
        <div class='metric-value' style='color:#F59E0B;'>{warning_count}</div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown(f"""
    <div class='metric-card'>
        <div class='metric-label'>Critical (>= {critical_threshold} Leaves)</div>
        <div class='metric-value' style='color:#EF4444;'>{critical_count}</div>
    </div>
    """, unsafe_allow_html=True)

with col4:
    mode_text = "Live (SMTP)" if not simulation_mode else "Simulation"
    st.markdown(f"""
    <div class='metric-card'>
        <div class='metric-label'>Mailing Status</div>
        <div class='metric-value'>{mode_text}</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# Section 2: Core Workspaces (Left: Database view, Right: Action form)
left_col, right_col = st.columns([7, 5])

with left_col:
    st.markdown("### Student Attendance Records")
    
    # Simple filters to explore data
    filter_col1, filter_col2 = st.columns(2)
    with filter_col1:
        search_query = st.text_input("Search student by email/roll", "")
    with filter_col2:
        subject_filter = st.selectbox("Subject Filter", ["All", "CI", "Python", "DM"])
        
    df_filtered = df_attendance.copy()
    
    # Search
    if search_query:
        df_filtered = df_filtered[
            df_filtered["Student Email"].astype(str).str.contains(search_query, case=False) | 
            df_filtered["Roll Number"].astype(str).str.contains(search_query)
        ]
        
    # Subject filter leaves
    if subject_filter != "All":
        df_filtered = df_filtered.sort_values(by=subject_filter, ascending=False)
        
    # Helper to style leave counts with color status
    def color_attendance(val):
        if val >= critical_threshold:
            return 'background-color: rgba(239, 68, 68, 0.2); color: #EF4444; font-weight: bold;'
        elif val == warning_threshold:
            return 'background-color: rgba(245, 158, 11, 0.2); color: #F59E0B; font-weight: bold;'
        return 'background-color: rgba(16, 185, 129, 0.1); color: #10B981;'

    # Apply custom pandas styling to show warning/critical ranges visually
    # Use .map() for compatibility with modern Pandas versions (2.1.0+)
    styled_df = df_filtered.style.map(
        color_attendance, 
        subset=["CI", "Python", "DM"]
    )
    
    st.dataframe(styled_df, use_container_width=True, height=220, hide_index=True)
    
    st.caption("Green: Safe | Orange: Warning Threshold | Red: Critical Lack of Attendance")
    
    # Chart visualization showing absences per student
    if not df_filtered.empty:
        st.markdown("### Absences by Student")
        
        # Prepare chart data: set Roll Number as index and select CI, Python, DM
        chart_df = df_filtered.copy()
        chart_df["Roll Number"] = "Roll " + chart_df["Roll Number"].astype(str)
        chart_data = chart_df.set_index("Roll Number")[["CI", "Python", "DM"]]
        
        # Render the grouped bar chart
        st.bar_chart(chart_data, use_container_width=True, height=250)

with right_col:
    st.markdown("### Log Today's Absentees")
    
    if not is_authenticated:
        st.warning("Access Denied: Please enter the correct Staff Password in the sidebar access gate to unlock the attendance recording form.")
    else:
        with st.form("absentee_form", clear_on_submit=True):
            # Subject Selector
            subject_selected = st.selectbox("Choose Subject", list(SUBJECTS.keys()))
            
            # Student Multi-select dropdown
            # Combine Roll Number and Email for clear selection
            student_options = [
                f"Roll {row['Roll Number']} - {row['Student Email']}" 
                for _, row in df_attendance.iterrows()
            ]
            
            absentees_selected = st.multiselect(
                "Select Absent Student(s)",
                options=student_options,
                help="Select all students who are absent in today's class."
            )
            
            submit_button = st.form_submit_button("Record Absences & Process Alerts", use_container_width=True)
            
            if submit_button:
                if not absentees_selected:
                    st.warning("Please select at least one student.")
                else:
                    # Parse out the roll numbers
                    roll_numbers = []
                    for opt in absentees_selected:
                        # Extract roll number from format: "Roll X - email"
                        roll_num = int(opt.split(" ")[1])
                        roll_numbers.append(roll_num)
                    
                    subject_info = SUBJECTS[subject_selected]
                    subj_col = subject_info["col"]
                    subj_code = subject_info["code"]
                    
                    # Fetch staff mail for this subject
                    staff_mail = staff_emails[subj_col]
                    
                    # Load fresh copy of database
                    df_current = load_data().dropna(subset=["Roll Number"])
                    df_current["Roll Number"] = df_current["Roll Number"].astype(int)
                    
                    alerts_triggered = []
                    action_time = datetime.now().strftime("%H:%M:%S")
                    
                    # Process each student
                    for roll in roll_numbers:
                        # Find student row
                        student_idx = df_current[df_current["Roll Number"] == roll].index
                        
                        if len(student_idx) > 0:
                            idx = student_idx[0]
                            student_email = df_current.loc[idx, "Student Email"]
                            
                            # Increment leave count
                            old_leaves = df_current.loc[idx, subj_col]
                            new_leaves = old_leaves + 1
                            df_current.loc[idx, subj_col] = new_leaves
                            
                            action_log_msg = f"Roll {roll} leave incremented: {old_leaves} to {new_leaves} in {subj_col}"
                            st.session_state.action_logs.insert(0, f"[{action_time}] {action_log_msg}")
                            save_logs(st.session_state.email_logs, st.session_state.action_logs)
                            
                            # Threshold Alert checks
                            # Warning Threshold logic (Leaves reaches exactly warning_threshold)
                            if new_leaves == warning_threshold:
                                subject_name_long = "Computer Intelligence" if subj_col == "CI" else ("Python" if subj_col == "Python" else "Data Mining")
                                warning_msg = f"Dear Student, our records indicate {new_leaves} absences in {subject_name_long}. Please ensure regular attendance going forward."
                                
                                success, msg_details = send_notification(
                                    to_email=student_email,
                                    subject="Attendance Warning Report",
                                    body=warning_msg
                                )
                                alerts_triggered.append(f"Warning: Sent alert to Student {student_email} ({subj_col}: {new_leaves} leaves)")
                            
                            # Critical Threshold logic (Leaves exceeds warning threshold / reaches critical_threshold or more)
                            elif new_leaves >= critical_threshold:
                                subject_name_long = "Computer Intelligence" if subj_col == "CI" else ("Python" if subj_col == "Python" else "Data Mining")
                                
                                # Mail to Student
                                student_crit_msg = f"Dear Student, our records indicate {new_leaves} absences in {subject_name_long}. Please ensure regular attendance going forward."
                                send_notification(
                                    to_email=student_email,
                                    subject="Attendance Threshold Notification",
                                    body=student_crit_msg
                                )
                                
                                # Mail to Staff
                                staff_crit_msg = f"Dear Professor, this is to notify you that the student with Roll Number {roll} ({student_email}) has reached {new_leaves} absences in your subject, {subject_name_long}."
                                send_notification(
                                    to_email=staff_mail,
                                    subject=f"Attendance Threshold Notification - Roll {roll}",
                                    body=staff_crit_msg
                                )
                                alerts_triggered.append(f"Critical Alert: Sent to Student & Staff for {student_email} ({subj_col}: {new_leaves} leaves)")
                    
                    # Save updated DataFrame
                    save_data(df_current)
                    
                    # Show success results
                    st.success(f"Successfully recorded absences for {len(roll_numbers)} student(s) in {subj_col}!")
                    
                    # Print any triggered notifications
                    for alert in alerts_triggered:
                        st.toast(alert)
                        
                    # Rerun Streamlit page to show updated table
                    st.rerun()

# Section 3: Logging & Activity Feeds
st.markdown("---")
log_col1, log_col2 = st.columns(2)

with log_col1:
    st.markdown("### Recent Action Logs")
    if not st.session_state.action_logs:
        st.info("No activities performed in this session yet.")
    else:
        # Clear button
        if st.button("Clear Action Logs"):
            st.session_state.action_logs = []
            save_logs(st.session_state.email_logs, st.session_state.action_logs)
            st.rerun()
            
        for action in st.session_state.action_logs[:8]:  # show last 8 logs
            st.markdown(f"<div class='log-box'>{action}</div>", unsafe_allow_html=True)

with log_col2:
    st.markdown("### Email Notifications Log")
    if not st.session_state.email_logs:
        st.info("No emails generated yet.")
    else:
        # Clear button
        if st.button("Clear Email Logs"):
            st.session_state.email_logs = []
            save_logs(st.session_state.email_logs, st.session_state.action_logs)
            st.rerun()
            
        for log in st.session_state.email_logs[:6]:  # show last 6 email records
            status_color = "green" if log["status"] == "Sent" else ("orange" if log["status"] == "Simulated" else "red")
            
            st.markdown(f"""
            <div class='email-preview'>
                <strong>Time:</strong> {log['timestamp']} | 
                <strong>Status:</strong> <span style="color:{status_color}; font-weight:bold;">{log['status']}</span><br>
                <strong>To:</strong> {log['to']}<br>
                <strong>Subject:</strong> {log['subject']}<br>
                <strong>Message:</strong><br>
                <div style="background:#f1f5f9; padding:8px; border-radius:4px; margin-top:5px; font-family:monospace; color:#0f172a; border:1px solid #e2e8f0;">
                    {log['body']}
                </div>
                <div style="font-size:0.8rem; color:#64748b; margin-top:5px;">
                    {log['details']}
                </div>
            </div>
            """, unsafe_allow_html=True)

# Footer Credit Block
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #64748b; font-size: 0.85rem; padding: 10px;'>"
    "Attendance Tracker | Designed & Developed by <strong>Rajesh Yadav</strong> | Build Date: June 2026"
    "</div>",
    unsafe_allow_html=True
)
