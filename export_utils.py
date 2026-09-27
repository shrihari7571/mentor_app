from __future__ import annotations

from io import BytesIO
from datetime import datetime
from typing import Iterable

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether
)
from reportlab.graphics.shapes import Drawing, String
from reportlab.graphics.charts.barcharts import VerticalBarChart


from db import get_conn

HEADER_FILL = "1F4E78"
SUB_FILL = "D9EAF7"
LIGHT_FILL = "F3F6FA"
WHITE = "FFFFFF"
THIN = Side(style="thin", color="D9D9D9")


def _query_df(sql: str, params: tuple = ()) -> pd.DataFrame:
    conn = get_conn()
    try:
        return pd.read_sql_query(sql, conn, params=params)
    finally:
        conn.close()


def attendance_df(roll_no: int | None = None) -> pd.DataFrame:
    where = "WHERE a.roll_no=?" if roll_no is not None else ""
    params = (roll_no,) if roll_no is not None else ()
    return _query_df(
        f"""
        SELECT a.roll_no, s.urn, s.full_name AS student_name,
               a.month, a.percentage,
               CASE
                 WHEN a.percentage IS NULL THEN 'Not entered'
                 WHEN a.percentage >= 75 THEN 'Good'
                 WHEN a.percentage >= 65 THEN 'Watch'
                 ELSE 'Defaulter'
               END AS attendance_status,
               COALESCE(a.remarks, '') AS remarks
        FROM attendance a
        JOIN students s ON s.roll_no=a.roll_no
        {where}
        ORDER BY a.roll_no, a.month
        """, params
    )


def technical_df(roll_no: int | None = None) -> pd.DataFrame:
    where = "WHERE s.roll_no=?" if roll_no is not None else ""
    params = (roll_no,) if roll_no is not None else ()
    return _query_df(
        f"""
        SELECT s.roll_no, s.urn, s.full_name AS student_name, s.cpi,
               COALESCE(p.hackerrank_url,'') AS hackerrank_url,
               COALESCE(p.hackerrank_problems,0) AS hackerrank_problems,
               COALESCE(p.hackerrank_level,'') AS hackerrank_level,
               COALESCE(p.github_url,'') AS github_url,
               COALESCE(p.github_repos,0) AS github_repos,
               COALESCE(p.github_projects,'') AS github_projects
        FROM students s
        LEFT JOIN progress_profile p ON p.roll_no=s.roll_no
        {where}
        ORDER BY s.roll_no
        """, params
    )


def certifications_df(roll_no: int | None = None) -> pd.DataFrame:
    where = "WHERE c.roll_no=?" if roll_no is not None else ""
    params = (roll_no,) if roll_no is not None else ()
    return _query_df(
        f"""
        SELECT c.id, c.roll_no, s.urn, s.full_name AS student_name,
               c.course_name, COALESCE(c.provider,'') AS provider,
               COALESCE(c.status,'') AS status,
               COALESCE(c.start_date,'') AS start_date,
               COALESCE(c.completion_date,'') AS completion_date,
               COALESCE(c.certificate_link,'') AS certificate_link,
               COALESCE(c.remarks,'') AS remarks
        FROM certifications c
        JOIN students s ON s.roll_no=c.roll_no
        {where}
        ORDER BY c.roll_no, COALESCE(c.completion_date,c.start_date,c.created_at) DESC
        """, params
    )


def placement_df(roll_no: int | None = None) -> pd.DataFrame:
    where = "WHERE s.roll_no=?" if roll_no is not None else ""
    params = (roll_no,) if roll_no is not None else ()
    return _query_df(
        f"""
        SELECT s.roll_no, s.urn, s.full_name AS student_name,
               COALESCE(p.resume_status,'') AS resume_status,
               COALESCE(p.resume_last_updated,'') AS resume_last_updated,
               COALESCE(p.placement_status,'') AS placement_status,
               COALESCE(p.placement_readiness,'') AS placement_readiness,
               COALESCE(p.target_companies,'') AS target_companies,
               COALESCE(p.aptitude_status,'') AS aptitude_status,
               COALESCE(p.technical_interview_status,'') AS technical_interview_status,
               COALESCE(p.communication_status,'') AS communication_status,
               COALESCE(p.placement_notes,'') AS placement_notes
        FROM students s
        LEFT JOIN progress_profile p ON p.roll_no=s.roll_no
        {where}
        ORDER BY s.roll_no
        """, params
    )


def meetings_df(roll_no: int | None = None) -> pd.DataFrame:
    where = "WHERE m.roll_no=?" if roll_no is not None else ""
    params = (roll_no,) if roll_no is not None else ()
    return _query_df(
        f"""
        SELECT m.roll_no, s.urn, s.full_name AS student_name,
               m.meeting_no, COALESCE(m.meeting_date,'') AS meeting_date,
               COALESCE(m.previous_target,'') AS previous_target,
               COALESCE(m.student_update,'') AS student_update,
               COALESCE(m.achievement_status,'') AS achievement_status,
               COALESCE(m.mentor_remark,'') AS mentor_remark,
               COALESCE(m.next_target,'') AS next_target,
               COALESCE(m.next_review_date,'') AS next_review_date
        FROM meetings m
        JOIN students s ON s.roll_no=m.roll_no
        {where}
        ORDER BY m.roll_no, m.meeting_no
        """, params
    )


def student_profile_df(roll_no: int) -> pd.DataFrame:
    return _query_df(
        """
        SELECT roll_no, urn, full_name AS student_name, contact_no, email, cpi,
               goals, interests_hobbies, strengths, weaknesses, response_status
        FROM students WHERE roll_no=?
        """, (roll_no,)
    )


def _safe(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return value


def _format_sheet(ws, freeze="A2", autofilter=True):
    ws.freeze_panes = freeze
    if autofilter and ws.max_row >= 1 and ws.max_column >= 1:
        ws.auto_filter.ref = ws.dimensions
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=HEADER_FILL)
        cell.font = Font(color=WHITE, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in ws.iter_rows():
        for cell in row:
            cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for col in range(1, ws.max_column + 1):
        max_len = 0
        for row in range(1, min(ws.max_row, 200) + 1):
            val = ws.cell(row=row, column=col).value
            if val is not None:
                max_len = max(max_len, min(len(str(val)), 45))
        ws.column_dimensions[get_column_letter(col)].width = max(11, min(max_len + 2, 38))


def _write_df(ws, df: pd.DataFrame):
    cols = list(df.columns)
    ws.append(cols)
    for row in df.itertuples(index=False, name=None):
        ws.append([_safe(v) for v in row])
    _format_sheet(ws)


def _add_summary_sheet(wb: Workbook):
    ws = wb.active
    ws.title = "Summary"
    df = _query_df(
        """
        SELECT s.roll_no, s.urn, s.full_name AS student_name, s.cpi, s.response_status,
               a.month AS latest_attendance_month, a.percentage AS latest_attendance,
               COALESCE(p.hackerrank_problems,0) AS hackerrank_problems,
               COALESCE(p.github_repos,0) AS github_repos,
               COALESCE(p.resume_status,'') AS resume_status,
               COALESCE(p.placement_status,'') AS placement_status,
               COALESCE(p.placement_readiness,'') AS placement_readiness,
               (SELECT COUNT(*) FROM certifications c WHERE c.roll_no=s.roll_no AND c.status='Completed') AS completed_certifications,
               (SELECT COUNT(*) FROM meetings m WHERE m.roll_no=s.roll_no) AS meetings_recorded
        FROM students s
        LEFT JOIN progress_profile p ON p.roll_no=s.roll_no
        LEFT JOIN attendance a ON a.id=(
            SELECT a2.id FROM attendance a2 WHERE a2.roll_no=s.roll_no ORDER BY a2.month DESC LIMIT 1
        )
        ORDER BY s.roll_no
        """
    )
    if not df.empty:
        df["attendance_status"] = df["latest_attendance"].apply(
            lambda x: "" if pd.isna(x) else ("Good" if x >= 75 else ("Watch" if x >= 65 else "Defaulter"))
        )
    _write_df(ws, df)


def batch_excel_bytes(section: str = "complete") -> bytes:
    wb = Workbook()
    if section == "complete":
        _add_summary_sheet(wb)
        datasets = [
            ("Attendance", attendance_df()),
            ("Technical Progress", technical_df()),
            ("Certifications", certifications_df()),
            ("Resume Placement", placement_df()),
            ("Meeting Progress", meetings_df()),
        ]
        for title, df in datasets:
            ws = wb.create_sheet(title)
            _write_df(ws, df)
    else:
        wb.remove(wb.active)
        mapping = {
            "attendance": ("Attendance", attendance_df()),
            "technical": ("Technical Progress", technical_df()),
            "certifications": ("Certifications", certifications_df()),
            "placement": ("Resume Placement", placement_df()),
            "meetings": ("Meeting Progress", meetings_df()),
        }
        title, df = mapping[section]
        ws = wb.create_sheet(title)
        _write_df(ws, df)
    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()


def student_excel_bytes(roll_no: int) -> bytes:
    profile = student_profile_df(roll_no)
    if profile.empty:
        raise ValueError(f"Student {roll_no} not found")
    wb = Workbook()
    ws = wb.active
    ws.title = "Student Profile"
    p = profile.iloc[0].to_dict()
    ws.append(["MENTOR-MENTEE INDIVIDUAL STUDENT REPORT", ""])
    ws.merge_cells("A1:B1")
    ws["A1"].fill = PatternFill("solid", fgColor=HEADER_FILL)
    ws["A1"].font = Font(color=WHITE, bold=True, size=14)
    ws["A1"].alignment = Alignment(horizontal="center")
    fields = [
        ("Roll No.", p.get("roll_no")), ("URN", p.get("urn")),
        ("Student Name", p.get("student_name")), ("CPI", p.get("cpi")),
        ("Contact", p.get("contact_no")), ("Email", p.get("email")),
        ("Response Status", p.get("response_status")), ("Goals", p.get("goals")),
        ("Interests & Hobbies", p.get("interests_hobbies")), ("Strengths", p.get("strengths")),
        ("Development Areas", p.get("weaknesses")),
    ]
    for label, value in fields:
        ws.append([label, _safe(value)])
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 70
    for row in ws.iter_rows(min_row=2):
        row[0].font = Font(bold=True)
        row[0].fill = PatternFill("solid", fgColor=SUB_FILL)
        for cell in row:
            cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    datasets = [
        ("Attendance", attendance_df(roll_no)),
        ("Technical Progress", technical_df(roll_no)),
        ("Certifications", certifications_df(roll_no)),
        ("Resume Placement", placement_df(roll_no)),
        ("Meeting Progress", meetings_df(roll_no)),
    ]
    for title, df in datasets:
        ws2 = wb.create_sheet(title)
        _write_df(ws2, df)

    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()


def _p(text, style):
    if text is None:
        val = ""
    else:
        try:
            val = "" if pd.isna(text) else str(text)
        except (TypeError, ValueError):
            val = str(text)
    val = val.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    val = val.replace("\n", "<br/>")
    return Paragraph(val or "-", style)


def _page_footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.grey)
    canvas.drawString(15 * mm, 8 * mm, "Private mentor report - generated locally")
    canvas.drawRightString(282 * mm, 8 * mm, f"Page {doc.page}")
    canvas.restoreState()


def student_pdf_bytes(roll_no: int) -> bytes:
    profile = student_profile_df(roll_no)
    if profile.empty:
        raise ValueError(f"Student {roll_no} not found")
    p = profile.iloc[0].to_dict()
    attendance = attendance_df(roll_no)
    technical = technical_df(roll_no)
    certs = certifications_df(roll_no)
    placement = placement_df(roll_no)
    meetings = meetings_df(roll_no)

    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(A4),
        rightMargin=12 * mm,
        leftMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=14 * mm,
        title=f"Mentor Report - {p.get('student_name')}",
        author="Mentor-Mentee Tracker",
    )
    styles = getSampleStyleSheet()
    title = ParagraphStyle(
        "ReportTitle", parent=styles["Title"], fontName="Helvetica-Bold",
        fontSize=18, leading=21, textColor=colors.HexColor("#1F4E78"), alignment=TA_CENTER,
        spaceAfter=8,
    )
    h2 = ParagraphStyle(
        "Section", parent=styles["Heading2"], fontName="Helvetica-Bold",
        fontSize=11, leading=14, textColor=colors.HexColor("#1F4E78"), spaceBefore=7, spaceAfter=5,
    )
    body = ParagraphStyle("BodySmall", parent=styles["BodyText"], fontSize=8.5, leading=11)
    small = ParagraphStyle("Tiny", parent=styles["BodyText"], fontSize=7.3, leading=9)

    story = [
        _p("Mentor-Mentee Individual Progress Report", title),
        _p(f"Generated: {datetime.now().strftime('%d-%m-%Y %H:%M')}", body),
        Spacer(1, 4),
    ]

    basic_data = [
        [_p("Roll No.", body), _p(p.get("roll_no"), body), _p("URN", body), _p(p.get("urn"), body),
         _p("CPI", body), _p(p.get("cpi"), body)],
        [_p("Student Name", body), _p(p.get("student_name"), body), _p("Contact", body), _p(p.get("contact_no"), body),
         _p("Response", body), _p(p.get("response_status"), body)],
        [_p("Email", body), _p(p.get("email"), body), _p("Goals", body), _p(p.get("goals"), body),
         _p("Strengths", body), _p(p.get("strengths"), body)],
        [_p("Interests / Hobbies", body), _p(p.get("interests_hobbies"), body), _p("Development Areas", body),
         _p(p.get("weaknesses"), body), "", ""],
    ]
    t = Table(basic_data, colWidths=[22*mm, 42*mm, 22*mm, 66*mm, 22*mm, 76*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (0,-1), colors.HexColor("#D9EAF7")),
        ("BACKGROUND", (2,0), (2,-1), colors.HexColor("#D9EAF7")),
        ("BACKGROUND", (4,0), (4,-1), colors.HexColor("#D9EAF7")),
        ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("GRID", (0,0), (-1,-1), 0.35, colors.HexColor("#CFCFCF")),
        ("LEFTPADDING", (0,0), (-1,-1), 4), ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("TOPPADDING", (0,0), (-1,-1), 4), ("BOTTOMPADDING", (0,0), (-1,-1), 4),
    ]))
    story += [t]

    story += [_p("Attendance", h2)]
    if attendance.empty:
        story += [_p("No attendance records entered.", body)]
    else:
        data = [[_p(x, small) for x in ["Month", "Attendance %", "Status", "Remarks"]]]
        for _, r in attendance.iterrows():
            data.append([_p(r["month"], small), _p(r["percentage"], small), _p(r["attendance_status"], small), _p(r["remarks"], small)])
        tb = Table(data, colWidths=[35*mm, 30*mm, 28*mm, 150*mm], repeatRows=1)
        tb.setStyle(_pdf_table_style())
        story += [tb]

    story += [_p("Technical Progress and Placement", h2)]
    tech = technical.iloc[0].to_dict() if not technical.empty else {}
    plc = placement.iloc[0].to_dict() if not placement.empty else {}
    tech_data = [
        [_p("HackerRank", body), _p(tech.get("hackerrank_url"), body), _p("Problems", body), _p(tech.get("hackerrank_problems"), body),
         _p("Level", body), _p(tech.get("hackerrank_level"), body)],
        [_p("GitHub", body), _p(tech.get("github_url"), body), _p("Repos", body), _p(tech.get("github_repos"), body),
         _p("Projects", body), _p(tech.get("github_projects"), body)],
        [_p("Resume", body), _p(plc.get("resume_status"), body), _p("Placement Status", body), _p(plc.get("placement_status"), body),
         _p("Readiness", body), _p(plc.get("placement_readiness"), body)],
        [_p("Target Companies", body), _p(plc.get("target_companies"), body), _p("Aptitude", body), _p(plc.get("aptitude_status"), body),
         _p("Technical Interview", body), _p(plc.get("technical_interview_status"), body)],
        [_p("Communication / HR", body), _p(plc.get("communication_status"), body), _p("Mentor Placement Notes", body), _p(plc.get("placement_notes"), body), "", ""],
    ]
    tt = Table(tech_data, colWidths=[28*mm, 68*mm, 27*mm, 34*mm, 31*mm, 66*mm])
    tt.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (0,-1), colors.HexColor("#D9EAF7")),
        ("BACKGROUND", (2,0), (2,-1), colors.HexColor("#D9EAF7")),
        ("BACKGROUND", (4,0), (4,-1), colors.HexColor("#D9EAF7")),
        ("GRID", (0,0), (-1,-1), 0.35, colors.HexColor("#CFCFCF")),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 4), ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("TOPPADDING", (0,0), (-1,-1), 4), ("BOTTOMPADDING", (0,0), (-1,-1), 4),
    ]))
    story += [tt]

    story += [_p("Certifications", h2)]
    if certs.empty:
        story += [_p("No certification records entered.", body)]
    else:
        data = [[_p(x, small) for x in ["Course", "Provider", "Status", "Start", "Completion", "Evidence / Remarks"]]]
        for _, r in certs.iterrows():
            evidence = " / ".join(x for x in [str(r["certificate_link"] or "").strip(), str(r["remarks"] or "").strip()] if x)
            data.append([_p(r["course_name"], small), _p(r["provider"], small), _p(r["status"], small), _p(r["start_date"], small), _p(r["completion_date"], small), _p(evidence, small)])
        tb = Table(data, colWidths=[60*mm, 42*mm, 28*mm, 26*mm, 28*mm, 76*mm], repeatRows=1)
        tb.setStyle(_pdf_table_style())
        story += [tb]

    story += [_p("Meeting-wise Mentoring Progress", h2)]
    if meetings.empty:
        story += [_p("No meeting-progress records entered.", body)]
    else:
        data = [[_p(x, small) for x in ["No.", "Date", "Previous Target", "Student Update / Evidence", "Status", "Mentor Remark", "Next Target", "Next Review"]]]
        for _, r in meetings.iterrows():
            data.append([
                _p(r["meeting_no"], small), _p(r["meeting_date"], small), _p(r["previous_target"], small),
                _p(r["student_update"], small), _p(r["achievement_status"], small), _p(r["mentor_remark"], small),
                _p(r["next_target"], small), _p(r["next_review_date"], small),
            ])
        tb = Table(data, colWidths=[10*mm, 22*mm, 42*mm, 45*mm, 21*mm, 42*mm, 45*mm, 27*mm], repeatRows=1)
        tb.setStyle(_pdf_table_style())
        story += [tb]

    story += [Spacer(1, 8), _p("This report intentionally excludes private medical and emergency details.", small)]
    doc.build(story, onFirstPage=_page_footer, onLaterPages=_page_footer)
    return buf.getvalue()


def _pdf_table_style():
    return TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1F4E78")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), 0.35, colors.HexColor("#CFCFCF")),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 3), ("RIGHTPADDING", (0,0), (-1,-1), 3),
        ("TOPPADDING", (0,0), (-1,-1), 3), ("BOTTOMPADDING", (0,0), (-1,-1), 3),
    ])


# ---------------- Institute-wide PDF summary ----------------

def institute_summary_df() -> pd.DataFrame:
    """One institute-safe mentoring summary row per student."""
    df = _query_df(
        """
        SELECT s.roll_no, s.urn, s.full_name AS student_name, s.cpi,
               a.month AS latest_attendance_month, a.percentage AS latest_attendance,
               COALESCE(p.hackerrank_problems,0) AS hackerrank_problems,
               COALESCE(p.github_repos,0) AS github_repos,
               COALESCE(p.resume_status,'') AS resume_status,
               COALESCE(p.placement_readiness,'') AS placement_readiness,
               COALESCE(p.placement_status,'') AS placement_status,
               (SELECT COUNT(*) FROM certifications c
                    WHERE c.roll_no=s.roll_no AND c.status='Completed') AS completed_certifications
        FROM students s
        LEFT JOIN progress_profile p ON p.roll_no=s.roll_no
        LEFT JOIN attendance a ON a.id=(
            SELECT a2.id FROM attendance a2
            WHERE a2.roll_no=s.roll_no ORDER BY a2.month DESC LIMIT 1
        )
        ORDER BY s.roll_no
        """
    )
    if df.empty:
        return df
    df["attendance_status"] = df["latest_attendance"].apply(
        lambda x: "Not entered" if pd.isna(x) else ("Good" if float(x) >= 75 else ("Watch" if float(x) >= 65 else "Defaulter"))
    )
    return df


def _count_status(series: pd.Series, values: list[str]) -> int:
    normalized = series.fillna("").astype(str).str.strip()
    return int(normalized.isin(values).sum())


def _bar_chart_drawing(title: str, labels: list[str], values: list[int], width=124*mm, height=62*mm):
    """Small self-contained dashboard-style bar chart for ReportLab PDF."""
    drawing = Drawing(width, height)
    drawing.add(String(width/2, height-10, title, textAnchor="middle", fontName="Helvetica-Bold", fontSize=10, fillColor=colors.HexColor("#1F4E78")))
    chart = VerticalBarChart()
    chart.x = 30
    chart.y = 24
    chart.width = width - 45
    chart.height = height - 48
    chart.data = [values]
    chart.categoryAxis.categoryNames = labels
    chart.categoryAxis.labels.fontName = "Helvetica"
    chart.categoryAxis.labels.fontSize = 6.5
    chart.categoryAxis.labels.angle = 20 if any(len(x) > 10 for x in labels) else 0
    chart.categoryAxis.labels.dy = -4
    chart.valueAxis.valueMin = 0
    max_val = max(values) if values else 0
    chart.valueAxis.valueMax = max(5, max_val + max(1, int(max_val * 0.2)))
    chart.valueAxis.valueStep = max(1, int(chart.valueAxis.valueMax / 5))
    chart.valueAxis.labels.fontSize = 7
    chart.bars[0].fillColor = colors.HexColor("#5B9BD5")
    chart.bars[0].strokeColor = colors.HexColor("#4472C4")
    chart.barLabelFormat = '%d'
    chart.barLabels.nudge = 5
    chart.barLabels.fontSize = 7
    drawing.add(chart)
    return drawing


def _institute_footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.grey)
    canvas.drawString(15 * mm, 8 * mm, "Institute mentor-mentee summary - generated locally")
    canvas.drawRightString(282 * mm, 8 * mm, f"Page {doc.page}")
    canvas.restoreState()


def institute_pdf_bytes() -> bytes:
    """Create a privacy-safe institute-wide mentoring summary with dashboard charts."""
    summary = institute_summary_df()
    tech = technical_df()
    certs = certifications_df()
    placement = placement_df()

    total = int(len(summary))
    att_counts = {
        "Good": int((summary["attendance_status"] == "Good").sum()) if total else 0,
        "Watch": int((summary["attendance_status"] == "Watch").sum()) if total else 0,
        "Defaulter": int((summary["attendance_status"] == "Defaulter").sum()) if total else 0,
        "Not entered": int((summary["attendance_status"] == "Not entered").sum()) if total else 0,
    }

    hr_profile = int(tech["hackerrank_url"].fillna("").astype(str).str.strip().ne("").sum()) if not tech.empty else 0
    gh_profile = int(tech["github_url"].fillna("").astype(str).str.strip().ne("").sum()) if not tech.empty else 0
    coding_active = int((pd.to_numeric(tech["hackerrank_problems"], errors="coerce").fillna(0) > 0).sum()) if not tech.empty else 0
    repo_active = int((pd.to_numeric(tech["github_repos"], errors="coerce").fillna(0) > 0).sum()) if not tech.empty else 0

    cert_status_order = ["Completed", "In Progress", "Planned", "Dropped"]
    cert_counts = [int((certs["status"].fillna("").astype(str).str.strip() == s).sum()) if not certs.empty else 0 for s in cert_status_order]

    resume_series = placement["resume_status"] if not placement.empty else pd.Series(dtype=str)
    resume_labels = ["Ready", "In Progress", "Not entered"]
    resume_values = [
        _count_status(resume_series, ["Completed", "Updated"]),
        _count_status(resume_series, ["In Progress"]),
        total - _count_status(resume_series, ["Completed", "Updated", "In Progress"]),
    ]

    readiness_series = placement["placement_readiness"] if not placement.empty else pd.Series(dtype=str)
    readiness_labels = ["Ready", "Moderate", "Basic", "Not entered"]
    readiness_values = [
        _count_status(readiness_series, ["Placement Ready"]),
        _count_status(readiness_series, ["Moderate"]),
        _count_status(readiness_series, ["Basic"]),
        total - _count_status(readiness_series, ["Placement Ready", "Moderate", "Basic"]),
    ]

    completed_cert_total = int(sum(cert_counts[:1]))
    resume_ready = int(resume_values[0])
    placement_ready = int(readiness_values[0])
    students_with_cert = int(certs.loc[certs["status"] == "Completed", "roll_no"].nunique()) if not certs.empty else 0

    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=landscape(A4),
        rightMargin=12*mm, leftMargin=12*mm, topMargin=11*mm, bottomMargin=14*mm,
        title="Institute Mentor-Mentee Summary Report", author="Mentor-Mentee Tracker",
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "InstituteTitle", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=18,
        leading=21, textColor=colors.HexColor("#1F4E78"), alignment=TA_CENTER, spaceAfter=6,
    )
    h2 = ParagraphStyle(
        "InstituteSection", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=11,
        leading=14, textColor=colors.HexColor("#1F4E78"), spaceBefore=6, spaceAfter=4,
    )
    body = ParagraphStyle("InstituteBody", parent=styles["BodyText"], fontSize=8.3, leading=10.5)
    tiny = ParagraphStyle("InstituteTiny", parent=styles["BodyText"], fontSize=6.6, leading=8)

    story = [
        _p("Mentor-Mentee Institute-Wide Progress Summary", title_style),
        _p(f"Generated: {datetime.now().strftime('%d-%m-%Y %H:%M')} | Cohort size: {total}", body),
        Spacer(1, 4),
    ]

    kpis = [
        ("Total mentees", total),
        ("Attendance >=75%", att_counts["Good"]),
        ("Attendance defaulters", att_counts["Defaulter"]),
        ("HackerRank profiles", hr_profile),
        ("GitHub profiles", gh_profile),
        ("Completed certifications", completed_cert_total),
        ("Students with completed cert.", students_with_cert),
        ("Resume ready", resume_ready),
        ("Placement ready", placement_ready),
    ]
    kpi_data = []
    row = []
    for label, value in kpis:
        cell = Table([
            [_p(label, tiny)],
            [_p(str(value), ParagraphStyle("KPIValue", parent=body, fontName="Helvetica-Bold", fontSize=14, leading=16, alignment=TA_CENTER, textColor=colors.HexColor("#1F4E78")))],
        ], colWidths=[28*mm])
        cell.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#F3F6FA")),
            ("BOX", (0,0), (-1,-1), 0.4, colors.HexColor("#D0D7DE")),
            ("ALIGN", (0,0), (-1,-1), "CENTER"),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("TOPPADDING", (0,0), (-1,-1), 4), ("BOTTOMPADDING", (0,0), (-1,-1), 4),
        ]))
        row.append(cell)
        if len(row) == 9:
            kpi_data.append(row)
    kpi_table = Table(kpi_data, colWidths=[29.5*mm]*9)
    story += [kpi_table, Spacer(1, 7)]

    story += [_p("Dashboard charts", h2)]
    attendance_chart = _bar_chart_drawing(
        "Attendance status",
        ["Good", "Watch", "Defaulter", "Not entered"],
        [att_counts["Good"], att_counts["Watch"], att_counts["Defaulter"], att_counts["Not entered"]],
    )
    tech_chart = _bar_chart_drawing(
        "Technical progress",
        ["HackerRank", "GitHub", "Coding active", "Repo active"],
        [hr_profile, gh_profile, coding_active, repo_active],
    )
    charts_top = Table([[attendance_chart, tech_chart]], colWidths=[132*mm, 132*mm])
    charts_top.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "TOP")]))
    story += [charts_top, Spacer(1, 5), PageBreak()]

    story += [_p("Professional readiness overview", h2)]
    cert_chart = _bar_chart_drawing("Certification status", cert_status_order, cert_counts, width=84*mm, height=66*mm)
    resume_chart = _bar_chart_drawing("Resume status", resume_labels, resume_values, width=84*mm, height=66*mm)
    placement_chart = _bar_chart_drawing("Placement readiness", readiness_labels, readiness_values, width=84*mm, height=66*mm)
    charts_bottom = Table([[cert_chart, resume_chart, placement_chart]], colWidths=[88*mm, 88*mm, 88*mm])
    charts_bottom.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "TOP")]))
    story += [charts_bottom, Spacer(1, 8)]

    story += [_p("Student-wise institute summary", h2)]
    table_header = ["Roll", "Student", "CPI", "Latest Att.%", "Att. Status", "HR Problems", "GitHub Repos", "Completed Cert.", "Resume", "Placement Readiness"]
    table_data = [[_p(x, tiny) for x in table_header]]
    for _, r in summary.iterrows():
        att = "" if pd.isna(r["latest_attendance"]) else f"{float(r['latest_attendance']):.1f}"
        table_data.append([
            _p(r["roll_no"], tiny), _p(r["student_name"], tiny), _p(r["cpi"], tiny), _p(att, tiny),
            _p(r["attendance_status"], tiny), _p(r["hackerrank_problems"], tiny), _p(r["github_repos"], tiny),
            _p(r["completed_certifications"], tiny), _p(r["resume_status"] or "Not entered", tiny),
            _p(r["placement_readiness"] or "Not entered", tiny),
        ])
    summary_table = Table(
        table_data,
        colWidths=[14*mm, 48*mm, 12*mm, 20*mm, 23*mm, 21*mm, 21*mm, 23*mm, 28*mm, 42*mm],
        repeatRows=1,
    )
    summary_table.setStyle(_pdf_table_style())
    story += [summary_table, Spacer(1, 5)]
    story += [_p(
        "Privacy: This institute-wide report excludes personal contact details, family information, addresses, medical information and emergency details. Blank values mean the item has not yet been entered in the tracker.",
        tiny,
    )]

    doc.build(story, onFirstPage=_institute_footer, onLaterPages=_institute_footer)
    return buf.getvalue()
