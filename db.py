from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Iterable, Optional

DB_PATH = Path(__file__).resolve().parent / "data" / "mentor.db"


def get_conn(db_path: Path | str = DB_PATH) -> sqlite3.Connection:
    db_path = Path(db_path)
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db(conn: Optional[sqlite3.Connection] = None) -> None:
    own = conn is None
    conn = conn or get_conn()
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS students (
            roll_no INTEGER PRIMARY KEY,
            urn TEXT UNIQUE NOT NULL,
            full_name TEXT NOT NULL,
            contact_no TEXT,
            email TEXT,
            cpi REAL,
            dob TEXT,
            gender TEXT,
            blood_group TEXT,
            father_name TEXT,
            residence_address TEXT,
            father_occupation TEXT,
            father_mobile TEXT,
            emergency_contact TEXT,
            goals TEXT,
            interests_hobbies TEXT,
            strengths TEXT,
            weaknesses TEXT,
            medical_condition TEXT,
            allergies TEXT,
            emergency_medication TEXT,
            extra_info TEXT,
            response_timestamp TEXT,
            response_status TEXT DEFAULT 'Pending',
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS meetings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            roll_no INTEGER NOT NULL,
            meeting_no INTEGER NOT NULL,
            meeting_date TEXT,
            previous_target TEXT,
            student_update TEXT,
            achievement_status TEXT,
            mentor_remark TEXT,
            next_target TEXT,
            next_review_date TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE (roll_no, meeting_no),
            FOREIGN KEY (roll_no) REFERENCES students(roll_no) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            roll_no INTEGER NOT NULL,
            month TEXT NOT NULL,
            percentage REAL,
            remarks TEXT,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE (roll_no, month),
            FOREIGN KEY (roll_no) REFERENCES students(roll_no) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS progress_profile (
            roll_no INTEGER PRIMARY KEY,
            hackerrank_url TEXT,
            hackerrank_problems INTEGER DEFAULT 0,
            hackerrank_level TEXT,
            github_url TEXT,
            github_repos INTEGER DEFAULT 0,
            github_projects TEXT,
            resume_status TEXT DEFAULT 'Not Started',
            resume_last_updated TEXT,
            placement_status TEXT DEFAULT 'Not Started',
            placement_readiness TEXT DEFAULT 'Not Started',
            target_companies TEXT,
            aptitude_status TEXT DEFAULT 'Not Started',
            technical_interview_status TEXT DEFAULT 'Not Started',
            communication_status TEXT DEFAULT 'Not Started',
            placement_notes TEXT,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (roll_no) REFERENCES students(roll_no) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS certifications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            roll_no INTEGER NOT NULL,
            course_name TEXT NOT NULL,
            provider TEXT,
            status TEXT DEFAULT 'In Progress',
            start_date TEXT,
            completion_date TEXT,
            certificate_link TEXT,
            remarks TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (roll_no) REFERENCES students(roll_no) ON DELETE CASCADE
        );

        CREATE INDEX IF NOT EXISTS idx_meetings_roll ON meetings(roll_no);
        CREATE INDEX IF NOT EXISTS idx_students_status ON students(response_status);
        CREATE INDEX IF NOT EXISTS idx_attendance_roll_month ON attendance(roll_no, month);
        CREATE INDEX IF NOT EXISTS idx_certifications_roll ON certifications(roll_no);
        """
    )

    # Ensure every existing student has a progress-profile row after upgrades.
    conn.execute(
        """
        INSERT OR IGNORE INTO progress_profile (roll_no)
        SELECT roll_no FROM students
        """
    )
    conn.commit()
    if own:
        conn.close()


def upsert_roster(rows: Iterable[tuple[int, str, str, str]], conn: Optional[sqlite3.Connection] = None) -> None:
    own = conn is None
    conn = conn or get_conn()
    conn.executemany(
        """
        INSERT INTO students (roll_no, urn, full_name, contact_no)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(roll_no) DO UPDATE SET
            urn=excluded.urn,
            full_name=excluded.full_name,
            contact_no=COALESCE(students.contact_no, excluded.contact_no)
        """,
        list(rows),
    )
    conn.execute(
        """
        INSERT OR IGNORE INTO progress_profile (roll_no)
        SELECT roll_no FROM students
        """
    )
    conn.commit()
    if own:
        conn.close()


def list_students(conn: Optional[sqlite3.Connection] = None):
    own = conn is None
    conn = conn or get_conn()
    rows = conn.execute("SELECT * FROM students ORDER BY roll_no").fetchall()
    if own:
        conn.close()
    return rows


def get_student(roll_no: int, conn: Optional[sqlite3.Connection] = None):
    own = conn is None
    conn = conn or get_conn()
    row = conn.execute("SELECT * FROM students WHERE roll_no=?", (roll_no,)).fetchone()
    if own:
        conn.close()
    return row


def save_student_notes(roll_no: int, goals: str, interests_hobbies: str, strengths: str, weaknesses: str,
                       conn: Optional[sqlite3.Connection] = None) -> None:
    own = conn is None
    conn = conn or get_conn()
    conn.execute(
        """
        UPDATE students SET goals=?, interests_hobbies=?, strengths=?, weaknesses=?,
        updated_at=CURRENT_TIMESTAMP WHERE roll_no=?
        """,
        (goals, interests_hobbies, strengths, weaknesses, roll_no),
    )
    conn.commit()
    if own:
        conn.close()


def get_meetings(roll_no: int, conn: Optional[sqlite3.Connection] = None):
    own = conn is None
    conn = conn or get_conn()
    rows = conn.execute(
        "SELECT * FROM meetings WHERE roll_no=? ORDER BY meeting_no", (roll_no,)
    ).fetchall()
    if own:
        conn.close()
    return rows


def upsert_meeting(data: dict, conn: Optional[sqlite3.Connection] = None) -> None:
    own = conn is None
    conn = conn or get_conn()
    conn.execute(
        """
        INSERT INTO meetings (
            roll_no, meeting_no, meeting_date, previous_target, student_update,
            achievement_status, mentor_remark, next_target, next_review_date
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(roll_no, meeting_no) DO UPDATE SET
            meeting_date=excluded.meeting_date,
            previous_target=excluded.previous_target,
            student_update=excluded.student_update,
            achievement_status=excluded.achievement_status,
            mentor_remark=excluded.mentor_remark,
            next_target=excluded.next_target,
            next_review_date=excluded.next_review_date,
            updated_at=CURRENT_TIMESTAMP
        """,
        (
            data["roll_no"], data["meeting_no"], data.get("meeting_date"),
            data.get("previous_target"), data.get("student_update"),
            data.get("achievement_status"), data.get("mentor_remark"),
            data.get("next_target"), data.get("next_review_date")
        ),
    )
    conn.commit()
    if own:
        conn.close()


def get_attendance(roll_no: int, conn: Optional[sqlite3.Connection] = None):
    own = conn is None
    conn = conn or get_conn()
    rows = conn.execute(
        "SELECT * FROM attendance WHERE roll_no=? ORDER BY month", (roll_no,)
    ).fetchall()
    if own:
        conn.close()
    return rows


def upsert_attendance(roll_no: int, month: str, percentage: float | None, remarks: str = "",
                      conn: Optional[sqlite3.Connection] = None) -> None:
    own = conn is None
    conn = conn or get_conn()
    conn.execute(
        """
        INSERT INTO attendance (roll_no, month, percentage, remarks)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(roll_no, month) DO UPDATE SET
            percentage=excluded.percentage,
            remarks=excluded.remarks,
            updated_at=CURRENT_TIMESTAMP
        """,
        (roll_no, month, percentage, remarks),
    )
    conn.commit()
    if own:
        conn.close()


def latest_attendance(roll_no: int, conn: Optional[sqlite3.Connection] = None):
    own = conn is None
    conn = conn or get_conn()
    row = conn.execute(
        "SELECT * FROM attendance WHERE roll_no=? ORDER BY month DESC LIMIT 1", (roll_no,)
    ).fetchone()
    if own:
        conn.close()
    return row


def get_progress_profile(roll_no: int, conn: Optional[sqlite3.Connection] = None):
    own = conn is None
    conn = conn or get_conn()
    conn.execute("INSERT OR IGNORE INTO progress_profile (roll_no) VALUES (?)", (roll_no,))
    conn.commit()
    row = conn.execute("SELECT * FROM progress_profile WHERE roll_no=?", (roll_no,)).fetchone()
    if own:
        conn.close()
    return row


def save_progress_profile(data: dict, conn: Optional[sqlite3.Connection] = None) -> None:
    own = conn is None
    conn = conn or get_conn()
    conn.execute(
        """
        INSERT INTO progress_profile (
            roll_no, hackerrank_url, hackerrank_problems, hackerrank_level,
            github_url, github_repos, github_projects, resume_status, resume_last_updated,
            placement_status, placement_readiness, target_companies, aptitude_status,
            technical_interview_status, communication_status, placement_notes
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(roll_no) DO UPDATE SET
            hackerrank_url=excluded.hackerrank_url,
            hackerrank_problems=excluded.hackerrank_problems,
            hackerrank_level=excluded.hackerrank_level,
            github_url=excluded.github_url,
            github_repos=excluded.github_repos,
            github_projects=excluded.github_projects,
            resume_status=excluded.resume_status,
            resume_last_updated=excluded.resume_last_updated,
            placement_status=excluded.placement_status,
            placement_readiness=excluded.placement_readiness,
            target_companies=excluded.target_companies,
            aptitude_status=excluded.aptitude_status,
            technical_interview_status=excluded.technical_interview_status,
            communication_status=excluded.communication_status,
            placement_notes=excluded.placement_notes,
            updated_at=CURRENT_TIMESTAMP
        """,
        (
            data["roll_no"], data.get("hackerrank_url"), int(data.get("hackerrank_problems") or 0),
            data.get("hackerrank_level"), data.get("github_url"), int(data.get("github_repos") or 0),
            data.get("github_projects"), data.get("resume_status"), data.get("resume_last_updated"),
            data.get("placement_status"), data.get("placement_readiness"), data.get("target_companies"),
            data.get("aptitude_status"), data.get("technical_interview_status"),
            data.get("communication_status"), data.get("placement_notes"),
        ),
    )
    conn.commit()
    if own:
        conn.close()


def get_certifications(roll_no: int, conn: Optional[sqlite3.Connection] = None):
    own = conn is None
    conn = conn or get_conn()
    rows = conn.execute(
        "SELECT * FROM certifications WHERE roll_no=? ORDER BY COALESCE(completion_date, start_date, created_at) DESC, id DESC",
        (roll_no,),
    ).fetchall()
    if own:
        conn.close()
    return rows


def add_certification(data: dict, conn: Optional[sqlite3.Connection] = None) -> None:
    own = conn is None
    conn = conn or get_conn()
    conn.execute(
        """
        INSERT INTO certifications (
            roll_no, course_name, provider, status, start_date, completion_date, certificate_link, remarks
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            data["roll_no"], data["course_name"], data.get("provider"), data.get("status"),
            data.get("start_date"), data.get("completion_date"), data.get("certificate_link"), data.get("remarks")
        ),
    )
    conn.commit()
    if own:
        conn.close()


def delete_certification(cert_id: int, conn: Optional[sqlite3.Connection] = None) -> None:
    own = conn is None
    conn = conn or get_conn()
    conn.execute("DELETE FROM certifications WHERE id=?", (cert_id,))
    conn.commit()
    if own:
        conn.close()


def dashboard_rows(conn: Optional[sqlite3.Connection] = None):
    own = conn is None
    conn = conn or get_conn()
    rows = conn.execute(
        """
        SELECT
            s.roll_no, s.urn, s.full_name, s.cpi, s.response_status,
            p.hackerrank_url, p.hackerrank_problems, p.hackerrank_level,
            p.github_url, p.github_repos, p.resume_status, p.placement_status,
            p.placement_readiness, p.target_companies,
            a.month AS attendance_month, a.percentage AS latest_attendance,
            (SELECT COUNT(*) FROM certifications c WHERE c.roll_no=s.roll_no AND c.status='Completed') AS certifications_completed
        FROM students s
        LEFT JOIN progress_profile p ON p.roll_no=s.roll_no
        LEFT JOIN attendance a ON a.id=(
            SELECT a2.id FROM attendance a2 WHERE a2.roll_no=s.roll_no ORDER BY a2.month DESC LIMIT 1
        )
        ORDER BY s.roll_no
        """
    ).fetchall()
    if own:
        conn.close()
    return rows
