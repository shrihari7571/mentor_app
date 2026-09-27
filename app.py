from __future__ import annotations

import shutil
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

from db import (
    DB_PATH,
    add_certification,
    dashboard_rows,
    delete_certification,
    get_attendance,
    get_certifications,
    get_conn,
    get_meetings,
    get_progress_profile,
    get_student,
    init_db,
    list_students,
    save_progress_profile,
    save_student_notes,
    upsert_attendance,
    upsert_meeting,
)
from importer import import_responses
from export_utils import batch_excel_bytes, student_excel_bytes, student_pdf_bytes, institute_pdf_bytes
from seed import seed

APP_DIR = Path(__file__).resolve().parent
BACKUP_DIR = APP_DIR / "backups"

st.set_page_config(page_title="Mentor–Mentee Tracker", page_icon="🎓", layout="wide")


def ensure_ready():
    if not DB_PATH.exists():
        seed()
    else:
        init_db()


def rows_to_df(rows):
    return pd.DataFrame([dict(r) for r in rows])


def student_label(row):
    return f"{row['roll_no']} – {row['full_name']}"


def status_index(options, value):
    value = value or options[0]
    return options.index(value) if value in options else 0


def attendance_status(value):
    if value is None or pd.isna(value):
        return "Not entered"
    value = float(value)
    if value >= 75:
        return "Good"
    if value >= 65:
        return "Watch"
    return "Defaulter"


ensure_ready()

st.sidebar.title("Mentor–Mentee")
page = st.sidebar.radio(
    "Navigation",
    [
        "Dashboard",
        "Student Profiles",
        "Attendance",
        "Technical & Placement",
        "Certifications",
        "Import Responses",
        "Meeting Progress",
        "Exports & Reports",
        "Backup",
    ],
)
st.sidebar.caption("Offline local app · SQLite database")

if page == "Dashboard":
    st.title("Mentor Dashboard")
    conn = get_conn()
    df = rows_to_df(dashboard_rows(conn))
    total = len(df)
    received = int((df["response_status"] == "Received").sum())
    pending = total - received
    latest_att = pd.to_numeric(df["latest_attendance"], errors="coerce")
    att_good = int((latest_att >= 75).sum())
    att_defaulters = int((latest_att < 65).sum())
    hr_profiles = int(df["hackerrank_url"].fillna("").astype(str).str.strip().ne("").sum())
    gh_profiles = int(df["github_url"].fillna("").astype(str).str.strip().ne("").sum())
    cert_completed = int(pd.to_numeric(df["certifications_completed"], errors="coerce").fillna(0).sum())
    resume_ready = int(df["resume_status"].isin(["Completed", "Updated"]).sum())
    placement_ready = int((df["placement_readiness"] == "Placement Ready").sum())

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Mentees", total)
    c2.metric("Responses Received", received)
    c3.metric("Latest Attendance ≥75%", att_good)
    c4.metric("Attendance Defaulters", att_defaulters)

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("HackerRank Profiles", hr_profiles)
    c6.metric("GitHub Profiles", gh_profiles)
    c7.metric("Completed Certifications", cert_completed)
    c8.metric("Resume Ready", resume_ready)

    c9, c10, c11 = st.columns(3)
    c9.metric("Placement Ready", placement_ready)
    c10.metric("Responses Pending", pending)
    c11.metric("CPI ≥ 8.0", int((pd.to_numeric(df["cpi"], errors="coerce") >= 8).sum()))

    st.subheader("Student progress summary")
    summary = df.copy()
    summary["Attendance Status"] = summary["latest_attendance"].apply(attendance_status)
    summary = summary[[
        "roll_no", "full_name", "cpi", "latest_attendance", "Attendance Status",
        "hackerrank_problems", "github_repos", "certifications_completed", "resume_status",
        "placement_readiness", "response_status"
    ]]
    summary.columns = [
        "Roll No.", "Student Name", "CPI", "Latest Attendance %", "Attendance Status",
        "HackerRank Problems", "GitHub Repos", "Certifications Completed", "Resume",
        "Placement Readiness", "Response"
    ]
    st.dataframe(summary, use_container_width=True, hide_index=True)

    left, right = st.columns(2)
    with left:
        st.subheader("Latest attendance overview")
        chart = df[["roll_no", "latest_attendance"]].dropna().set_index("roll_no")
        if not chart.empty:
            st.bar_chart(chart)
        else:
            st.info("Enter monthly attendance to see this chart.")
    with right:
        st.subheader("HackerRank problems solved")
        chart = df[["roll_no", "hackerrank_problems"]].copy()
        chart["hackerrank_problems"] = pd.to_numeric(chart["hackerrank_problems"], errors="coerce").fillna(0)
        chart = chart.set_index("roll_no")
        st.bar_chart(chart)

    if att_defaulters:
        rolls = df.loc[latest_att < 65, "roll_no"].tolist()
        st.warning("Attendance follow-up required: " + ", ".join(map(str, rolls)))
    conn.close()

elif page == "Student Profiles":
    st.title("Student Profiles")
    students = list_students()
    selected = st.selectbox("Select mentee", students, format_func=student_label)
    s = dict(get_student(selected["roll_no"]))
    p = dict(get_progress_profile(s["roll_no"]))
    latest_att_rows = get_attendance(s["roll_no"])
    latest_att = dict(latest_att_rows[-1]) if latest_att_rows else None
    certs = get_certifications(s["roll_no"])

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Roll No.", s["roll_no"])
    c2.metric("URN", s["urn"])
    c3.metric("CPI", "–" if s["cpi"] is None else s["cpi"])
    c4.metric("Attendance", "–" if not latest_att or latest_att["percentage"] is None else f"{latest_att['percentage']:.1f}%")
    c5.metric("Response", s["response_status"])

    left, right = st.columns(2)
    with left:
        st.subheader("Basic information")
        st.write(f"**Name:** {s['full_name']}")
        st.write(f"**Email:** {s['email'] or '—'}")
        st.write(f"**Mobile:** {s['contact_no'] or '—'}")
        st.write(f"**DOB:** {s['dob'] or '—'}")
        st.write(f"**Gender:** {s['gender'] or '—'}")
        st.write(f"**Blood group:** {s['blood_group'] or '—'}")

        st.subheader("Career / development")
        goals = st.text_area("Goals", value=s["goals"] or "", height=130)
        interests = st.text_area("Areas of interest & hobbies", value=s["interests_hobbies"] or "", height=100)
        strengths = st.text_area("Strengths", value=s["strengths"] or "", height=100)
        weaknesses = st.text_area("Weaknesses / development areas", value=s["weaknesses"] or "", height=100)
        if st.button("Save profile notes", type="primary"):
            save_student_notes(s["roll_no"], goals, interests, strengths, weaknesses)
            st.success("Profile notes saved.")

    with right:
        st.subheader("Current technical / placement snapshot")
        st.write(f"**HackerRank:** {p['hackerrank_url'] or '—'}")
        st.write(f"**Problems solved:** {p['hackerrank_problems'] or 0}")
        st.write(f"**GitHub:** {p['github_url'] or '—'}")
        st.write(f"**Repositories:** {p['github_repos'] or 0}")
        st.write(f"**Resume:** {p['resume_status'] or '—'}")
        st.write(f"**Placement readiness:** {p['placement_readiness'] or '—'}")
        st.write(f"**Completed certifications:** {sum(1 for c in certs if c['status'] == 'Completed')}")

        st.subheader("Family & emergency information")
        st.write(f"**Father's name:** {s['father_name'] or '—'}")
        st.write(f"**Occupation:** {s['father_occupation'] or '—'}")
        st.write(f"**Father mobile:** {s['father_mobile'] or '—'}")
        st.write(f"**Emergency contact:** {s['emergency_contact'] or '—'}")
        st.write(f"**Address:** {s['residence_address'] or '—'}")

        with st.expander("Private health / emergency details"):
            st.write(f"**Medical condition:** {s['medical_condition'] or '—'}")
            st.write(f"**Allergies:** {s['allergies'] or '—'}")
            st.write(f"**Emergency medication:** {s['emergency_medication'] or '—'}")
            st.write(f"**Other information:** {s['extra_info'] or '—'}")

    st.subheader("Meeting history")
    meetings = rows_to_df(get_meetings(s["roll_no"]))
    if meetings.empty:
        st.info("No meeting progress recorded yet.")
    else:
        cols = ["meeting_no", "meeting_date", "achievement_status", "next_target", "next_review_date"]
        st.dataframe(meetings[cols], use_container_width=True, hide_index=True)

elif page == "Attendance":
    st.title("Monthly Attendance Tracking")
    students = list_students()
    selected = st.selectbox("Select mentee", students, format_func=student_label, key="att_student")
    roll = selected["roll_no"]
    history = rows_to_df(get_attendance(roll))

    c1, c2, c3 = st.columns(3)
    with c1:
        year = st.number_input("Year", min_value=2024, max_value=2035, value=date.today().year, step=1)
    with c2:
        month_no = st.selectbox("Month", list(range(1, 13)), index=max(0, date.today().month - 1), format_func=lambda m: pd.Timestamp(2000, m, 1).strftime("%B"))
    with c3:
        percentage = st.number_input("Attendance %", min_value=0.0, max_value=100.0, value=75.0, step=0.1)
    remarks = st.text_input("Remarks / reason for low attendance")

    if st.button("Save attendance", type="primary"):
        upsert_attendance(roll, f"{int(year):04d}-{int(month_no):02d}", float(percentage), remarks)
        st.success("Attendance saved.")
        st.rerun()

    st.subheader("Attendance history")
    history = rows_to_df(get_attendance(roll))
    if history.empty:
        st.info("No attendance entered yet.")
    else:
        history["Status"] = history["percentage"].apply(attendance_status)
        st.dataframe(history[["month", "percentage", "Status", "remarks"]], use_container_width=True, hide_index=True)
        st.line_chart(history.set_index("month")[["percentage"]])

elif page == "Technical & Placement":
    st.title("HackerRank, GitHub, Resume & Placement Tracking")
    students = list_students()
    selected = st.selectbox("Select mentee", students, format_func=student_label, key="tech_student")
    roll = selected["roll_no"]
    p = dict(get_progress_profile(roll))

    st.subheader("Technical profile")
    c1, c2 = st.columns(2)
    with c1:
        hackerrank_url = st.text_input("HackerRank profile link", value=p["hackerrank_url"] or "")
        hackerrank_problems = st.number_input("HackerRank problems solved", min_value=0, value=int(p["hackerrank_problems"] or 0), step=1)
        hr_options = ["Not Started", "Beginner", "Intermediate", "Advanced"]
        hackerrank_level = st.selectbox("HackerRank / coding level", hr_options, index=status_index(hr_options, p["hackerrank_level"] or "Not Started"))
    with c2:
        github_url = st.text_input("GitHub profile link", value=p["github_url"] or "")
        github_repos = st.number_input("GitHub repositories / projects", min_value=0, value=int(p["github_repos"] or 0), step=1)
        github_projects = st.text_area("Important GitHub projects / repositories", value=p["github_projects"] or "", height=90)

    st.subheader("Resume & placement")
    c3, c4 = st.columns(2)
    with c3:
        resume_options = ["Not Started", "In Progress", "Completed", "Updated"]
        resume_status = st.selectbox("Resume status", resume_options, index=status_index(resume_options, p["resume_status"] or "Not Started"))
        resume_last_updated = st.date_input(
            "Resume last updated",
            value=pd.to_datetime(p["resume_last_updated"]).date() if p["resume_last_updated"] else date.today(),
        )
        placement_status_options = ["Not Started", "Preparing", "Applied", "Shortlisted", "Interviewing", "Placed", "Higher Studies"]
        placement_status = st.selectbox("Placement status", placement_status_options, index=status_index(placement_status_options, p["placement_status"] or "Not Started"))
        readiness_options = ["Not Started", "Basic", "Moderate", "Placement Ready"]
        placement_readiness = st.selectbox("Placement readiness", readiness_options, index=status_index(readiness_options, p["placement_readiness"] or "Not Started"))
    with c4:
        target_companies = st.text_area("Target companies / organizations", value=p["target_companies"] or "", height=90)
        prep_options = ["Not Started", "Basic", "In Progress", "Good", "Ready"]
        aptitude_status = st.selectbox("Aptitude preparation", prep_options, index=status_index(prep_options, p["aptitude_status"] or "Not Started"))
        technical_interview_status = st.selectbox("Technical interview preparation", prep_options, index=status_index(prep_options, p["technical_interview_status"] or "Not Started"))
        communication_status = st.selectbox("Communication / HR preparation", prep_options, index=status_index(prep_options, p["communication_status"] or "Not Started"))

    placement_notes = st.text_area("Placement notes / mentor observations", value=p["placement_notes"] or "", height=100)

    if st.button("Save technical & placement progress", type="primary"):
        save_progress_profile({
            "roll_no": roll,
            "hackerrank_url": hackerrank_url,
            "hackerrank_problems": hackerrank_problems,
            "hackerrank_level": hackerrank_level,
            "github_url": github_url,
            "github_repos": github_repos,
            "github_projects": github_projects,
            "resume_status": resume_status,
            "resume_last_updated": resume_last_updated.isoformat(),
            "placement_status": placement_status,
            "placement_readiness": placement_readiness,
            "target_companies": target_companies,
            "aptitude_status": aptitude_status,
            "technical_interview_status": technical_interview_status,
            "communication_status": communication_status,
            "placement_notes": placement_notes,
        })
        st.success("Technical and placement progress saved.")

elif page == "Certifications":
    st.title("Certification Tracking")
    students = list_students()
    selected = st.selectbox("Select mentee", students, format_func=student_label, key="cert_student")
    roll = selected["roll_no"]

    with st.form("cert_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            course_name = st.text_input("Course / certification name *")
            provider = st.text_input("Provider (NPTEL / SWAYAM / Coursera / etc.)")
            status = st.selectbox("Status", ["Planned", "In Progress", "Completed", "Dropped"])
            start_date = st.date_input("Start date", value=date.today())
        with c2:
            completion_date = st.date_input("Completion date", value=date.today())
            certificate_link = st.text_input("Certificate / evidence link")
            remarks = st.text_area("Remarks", height=90)
        submitted = st.form_submit_button("Add certification", type="primary")
        if submitted:
            if not course_name.strip():
                st.error("Course / certification name is required.")
            else:
                add_certification({
                    "roll_no": roll,
                    "course_name": course_name.strip(),
                    "provider": provider.strip(),
                    "status": status,
                    "start_date": start_date.isoformat(),
                    "completion_date": completion_date.isoformat() if status == "Completed" else None,
                    "certificate_link": certificate_link.strip(),
                    "remarks": remarks.strip(),
                })
                st.success("Certification added.")
                st.rerun()

    certs = rows_to_df(get_certifications(roll))
    st.subheader("Certification history")
    if certs.empty:
        st.info("No certifications recorded yet.")
    else:
        st.dataframe(certs[["id", "course_name", "provider", "status", "start_date", "completion_date", "certificate_link", "remarks"]], use_container_width=True, hide_index=True)
        cert_id = st.selectbox("Delete a certification record (optional)", certs["id"].tolist(), format_func=lambda x: f"#{x} – {certs.loc[certs['id'] == x, 'course_name'].iloc[0]}")
        if st.button("Delete selected certification"):
            delete_certification(int(cert_id))
            st.success("Certification deleted.")
            st.rerun()

elif page == "Import Responses":
    st.title("Import Mentee Response Workbook")
    st.write("Upload the Excel file exported from your Google Form. Responses are matched by **Roll Number**; if a student entered the **URN** in that field, the app also maps it correctly.")
    uploaded = st.file_uploader("Choose .xlsx file", type=["xlsx"])
    if uploaded:
        if st.button("Import / Update Responses", type="primary"):
            try:
                result = import_responses(uploaded.getvalue())
                st.success(f"Imported/updated {result['mapped']} unique student responses.")
                if result["duplicates"]:
                    st.info("Duplicate responses detected; latest response retained for roll no.: " + ", ".join(map(str, result["duplicates"])))
                if result["pending"]:
                    st.warning("No response currently recorded for: " + ", ".join(map(str, result["pending"])))
                if result["unmatched"]:
                    st.error("Unmatched form entries: " + ", ".join(result["unmatched"]))
            except Exception as exc:
                st.exception(exc)

elif page == "Meeting Progress":
    st.title("Meeting-wise Progress Tracking")
    students = list_students()
    selected = st.selectbox("Select mentee", students, format_func=student_label, key="meeting_student")
    roll = selected["roll_no"]
    existing = {int(r["meeting_no"]): dict(r) for r in get_meetings(roll)}

    meeting_no = st.number_input("Meeting number", min_value=1, max_value=20, value=1, step=1)
    current = existing.get(int(meeting_no), {})

    c1, c2 = st.columns(2)
    with c1:
        meeting_date = st.date_input("Meeting date", value=pd.to_datetime(current.get("meeting_date")).date() if current.get("meeting_date") else date.today())
        previous_target = st.text_area("Previous target / action", value=current.get("previous_target") or "", height=100)
        student_update = st.text_area("Student update / evidence", value=current.get("student_update") or "", height=100)
        achievement = st.selectbox(
            "Achievement status",
            ["", "Yes", "Partially", "No", "Not Applicable"],
            index=["", "Yes", "Partially", "No", "Not Applicable"].index(current.get("achievement_status") or "") if (current.get("achievement_status") or "") in ["", "Yes", "Partially", "No", "Not Applicable"] else 0,
        )
    with c2:
        mentor_remark = st.text_area("Mentor remark", value=current.get("mentor_remark") or "", height=100)
        next_target = st.text_area("Next target / action", value=current.get("next_target") or "", height=100)
        next_review = st.date_input(
            "Next review date",
            value=pd.to_datetime(current.get("next_review_date")).date() if current.get("next_review_date") else date.today(),
        )

    if st.button("Save meeting progress", type="primary"):
        upsert_meeting({
            "roll_no": roll,
            "meeting_no": int(meeting_no),
            "meeting_date": meeting_date.isoformat(),
            "previous_target": previous_target,
            "student_update": student_update,
            "achievement_status": achievement,
            "mentor_remark": mentor_remark,
            "next_target": next_target,
            "next_review_date": next_review.isoformat(),
        })
        st.success("Meeting progress saved.")

    st.subheader("Recorded meetings")
    hist = rows_to_df(get_meetings(roll))
    if not hist.empty:
        st.dataframe(hist[["meeting_no", "meeting_date", "achievement_status", "mentor_remark", "next_target", "next_review_date"]], use_container_width=True, hide_index=True)
    else:
        st.info("No meeting records yet.")

elif page == "Exports & Reports":
    st.title("Exports & Mentor Reports")
    st.caption("Generate institute-ready Excel exports and private individual mentor reports. Medical and emergency details are excluded from reports by default.")

    st.subheader("Institute-wide PDF summary")
    st.write("Generate a privacy-safe institute summary with attendance, technical progress, certifications, resume status, placement readiness, dashboard charts, and a student-wise cohort table.")
    stamp = pd.Timestamp.now().strftime("%Y%m%d")
    st.download_button(
        "Download institute-wide PDF summary",
        data=institute_pdf_bytes(),
        file_name=f"Institute_Mentor_Mentee_Summary_{stamp}.pdf",
        mime="application/pdf",
        type="primary",
        use_container_width=True,
    )
    st.caption("Excludes student contact, family, address, medical and emergency details.")

    st.divider()
    st.subheader("Batch Excel exports")
    st.write("Download the complete batch workbook, or export only one tracking area.")

    c1, c2, c3 = st.columns(3)
    with c1:
        st.download_button(
            "Download complete Excel workbook",
            data=batch_excel_bytes("complete"),
            file_name=f"Mentor_Mentee_Complete_Export_{stamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        st.download_button(
            "Attendance Excel",
            data=batch_excel_bytes("attendance"),
            file_name=f"Attendance_Export_{stamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with c2:
        st.download_button(
            "Technical progress Excel",
            data=batch_excel_bytes("technical"),
            file_name=f"Technical_Progress_Export_{stamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.download_button(
            "Certifications Excel",
            data=batch_excel_bytes("certifications"),
            file_name=f"Certifications_Export_{stamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with c3:
        st.download_button(
            "Resume & placement Excel",
            data=batch_excel_bytes("placement"),
            file_name=f"Resume_Placement_Export_{stamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.download_button(
            "Meeting progress Excel",
            data=batch_excel_bytes("meetings"),
            file_name=f"Meeting_Progress_Export_{stamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.divider()
    st.subheader("Individual student mentor report")
    students = list_students()
    selected = st.selectbox("Select mentee for report", students, format_func=student_label, key="report_student")
    roll = int(selected["roll_no"])
    safe_name = "_".join(str(selected["full_name"]).strip().split())

    r1, r2 = st.columns(2)
    with r1:
        st.download_button(
            "Download individual Excel report",
            data=student_excel_bytes(roll),
            file_name=f"Mentor_Report_{roll}_{safe_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        st.caption("Includes student profile, attendance, technical progress, certifications, placement and meeting-wise progress.")
    with r2:
        st.download_button(
            "Download individual PDF mentor report",
            data=student_pdf_bytes(roll),
            file_name=f"Mentor_Report_{roll}_{safe_name}.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=True,
        )
        st.caption("Formatted multi-page PDF suitable for mentor files and institutional review.")

    st.info("Privacy note: exported mentor reports exclude medical conditions, allergies, emergency medication and emergency contact details. Keep exported files in a private mentor folder.")

elif page == "Backup":
    st.title("Backup")
    st.write("The entire application data is stored in one local SQLite file.")
    st.code(str(DB_PATH))
    BACKUP_DIR.mkdir(exist_ok=True)
    if st.button("Create backup now", type="primary"):
        stamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
        target = BACKUP_DIR / f"mentor_backup_{stamp}.db"
        shutil.copy2(DB_PATH, target)
        st.success(f"Backup created: {target.name}")
    backups = sorted(BACKUP_DIR.glob("*.db"), reverse=True)
    if backups:
        st.subheader("Available backups")
        for b in backups[:10]:
            st.write(b.name)
