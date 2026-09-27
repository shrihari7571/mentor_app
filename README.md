# Mentor-Mentee Tracker (Offline Streamlit + SQLite) - Version 4

This is a local/offline mentor-mentee tracking application for the 27-student batch.

## Included

- Mentor dashboard
- 27 preloaded mentees
- Student profile view
- Google Form Excel response import
- Roll Number / URN mapping and duplicate-response handling
- Monthly attendance tracking with history and trend chart
- HackerRank profile, problems solved and coding-level tracking
- GitHub profile, repository count and project notes
- Certification history (NPTEL / SWAYAM / Coursera / others)
- Resume-status tracking
- Placement status and placement-readiness tracking
- Aptitude, technical interview and communication/HR preparation status
- Target-company tracking
- Meeting-wise progress tracking
- **Export Center for Excel and PDF reports**
- Local SQLite database
- Local backup button

## Version 4 exports

Open **Exports & Reports** in the sidebar.

### Institute-wide PDF summary

Download a privacy-safe landscape PDF containing:

- Cohort KPI cards
- Attendance-status chart
- Technical-progress chart for HackerRank/GitHub activity
- Certification-status chart
- Resume-status chart
- Placement-readiness chart
- Student-wise institute summary table

The institute-wide report excludes student contact, family, address, medical and emergency information.

### Batch Excel exports

You can download:

- Complete mentor-mentee workbook
- Attendance workbook
- Technical progress workbook (HackerRank + GitHub)
- Certifications workbook
- Resume and placement workbook
- Meeting-progress workbook

The complete workbook contains separate sheets for:

1. Summary
2. Attendance
3. Technical Progress
4. Certifications
5. Resume Placement
6. Meeting Progress

### Individual student Excel report

Select a mentee and download an `.xlsx` report containing:

- Student profile and CPI
- Goals, interests, strengths and development areas
- Attendance history
- HackerRank and GitHub progress
- Certification history
- Resume and placement preparation
- Meeting-wise mentoring progress

### Individual student PDF mentor report

Select a mentee and download a formatted PDF report suitable for a mentor file or institutional review. The PDF includes:

- Basic academic/profile summary
- Attendance
- Technical progress
- Certifications
- Resume and placement preparation
- Meeting-wise targets, evidence, mentor remarks and next actions

For privacy, individual exported reports **do not include medical conditions, allergies, emergency medication or emergency-contact details**.

## Privacy

The app runs locally on the mentor's computer. Keep the project folder, `data/mentor.db`, backups, and exported reports private because the database contains personal information.

## First-time setup on Windows

1. Install Python 3.11 or 3.12 from python.org if Python is not already installed.
2. Extract this project folder to a private location, e.g. `C:\MentorTracker`.
3. Double-click `setup_windows.bat`.
4. Windows may ask for firewall permission; for local-only use, you can deny public-network access.
5. The app opens in your browser, normally at `http://127.0.0.1:8501`.

The first package installation may require internet. After the required Python packages are installed, normal use and report generation are local/offline.

## Later use

Double-click `run_app.bat`.

## Navigation

### Dashboard
Shows batch-level KPIs for latest attendance, HackerRank/GitHub adoption, completed certifications, resume readiness and placement readiness.

### Student Profiles
Shows personal information, latest attendance, technical/placement snapshot, completed-certification count, and meeting history.

### Attendance
Enter one monthly attendance percentage per student. Re-entering the same month updates that month's record instead of creating a duplicate.

### Technical & Placement
Track HackerRank, GitHub, resume, placement status, readiness, target companies, aptitude, technical-interview and communication/HR preparation.

### Certifications
Maintain multiple certification records per student with provider, status, dates, evidence link and remarks.

### Import Responses
Upload the `.xlsx` file exported from the Google Form. The importer matches Roll Number or URN, keeps the latest duplicate response, and lists pending students.

### Meeting Progress
Record mentor-mentee targets, evidence, status, mentor remarks and next review dates.

### Exports & Reports
Create batch Excel exports and individual Excel/PDF mentor reports directly in the browser. Files are generated in memory and downloaded to the location chosen by your browser.

### Backup
Create timestamped local backups of the SQLite database.

## Database

Data is stored in:

`data/mentor.db`

Version 4 uses the same Version 2 database structure. Existing student response, attendance, technical, certification, placement and meeting data is retained.

## Stop the app

Close the command window that launched Streamlit, or press `Ctrl+C` in that window.
