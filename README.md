# Attendance Tracker

A modern, interactive, and data-science-powered **Student Attendance Tracker & Email Reminder System** built using Python, Streamlit, and Pandas. This system allows academic staff to log student absences, view real-time metrics, analyze leave charts, and automate professional email warnings when students reach leave thresholds.

---

## Key Features

* **Excel Database Persistence**: All student leaves are loaded from and saved to a standard Excel spreadsheet (`attendance.xlsx`).
* **Interactive Data Table**: Color-coded records (Green: Safe, Orange: Warning, Red: Critical) with custom filtering and searching.
* **Interactive Data Visualization**: Grouped bar charts visualizing absences per student across all subjects.
* **Automatic Email Reminders**:
  * **Warning (Exactly 2 Absences)**: Sends a formal warning email to the student.
  * **Critical (3+ Absences)**: Sends formal lack-of-attendance notifications to both the student and the respective subject teacher.
* **Mailing Modes**:
  * **Simulation Mode (Default)**: Logs and displays emails on-screen for easy local testing.
  * **Live Mode**: Uses Gmail SMTP to send actual emails.
* **Staff Login/Sign Up**: Security gate preventing unauthorized editing. Credentials are saved in a local JSON database (`staff_credentials.json`).
* **Persistent Logs**: Action logs and email transmission history are saved to disk (`action_logs.json` and `email_logs.json`) so they persist across page refreshes.

---

## Tech Stack

* **Frontend Dashboard**: Streamlit (Python)
* **Data Management**: Pandas
* **Excel Engine**: openpyxl
* **Email Transmission**: smtplib & email (Python MIME API)
* **Storage**: Excel (for attendance) & JSON (for accounts and logs)

---

## File Structure

```
├── .streamlit/
│   └── config.toml          # Custom theme configuration (slate & emerald)
├── app.py                   # Main Streamlit web application
├── initialize_db.py         # Script to generate sample attendance Excel database
├── requirements.txt         # Project package dependencies
├── .gitignore               # Ignored logs and credential files
└── README.md                # Project documentation
```

---

## Setup and Installation

Follow these steps to run the application locally on your machine:

### 1. Clone the Repository
```bash
git clone https://github.com/Rajeshydv07/Attendance-Tracker.git
cd Attendance-Tracker
```

### 2. Install Dependencies
Make sure Python 3.8+ is installed, then run:
```bash
pip install -r requirements.txt
```

### 3. Initialize the Excel Spreadsheet
Run the helper script to create the initial `attendance.xlsx` template with sample student records:
```bash
python initialize_db.py
```

### 4. Run the Streamlit Application
Start the local web server:
```bash
python -m streamlit run app.py
```
This will compile and launch the dashboard in your web browser at `http://localhost:8501`.

---

## How to Use

1. **Log In**: Enter the default staff credentials in the **Staff Login** sidebar panel:
   * **Username**: `admin`
   * **Password**: `admin123`
   * *(Alternatively, switch to the "Sign Up" tab to register a new account).*
2. **View Records**: Inspect the student table and search/filter by name or subject.
3. **Log Absences**: Select a subject, select the absent student(s) from the dropdown list, and submit the form to update counts and trigger warnings.
4. **Inspect Emails**: Check the **Email Notifications Log** at the bottom of the page to review generated email warnings.

---

## Credits

* **Designed & Developed by**: Rajesh Yadav
* **Course/Department**: B.Tech - Information Technology
* **College**: ABES Engineering College
* **Build Date**: June 20, 2026
