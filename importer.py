from __future__ import annotations

from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import BinaryIO

import pandas as pd

from db import get_conn, init_db

REQUIRED_HEADERS = ["Full Name", "Roll Number"]


def _find_col(df: pd.DataFrame, target: str):
    target_norm = " ".join(target.strip().lower().split())
    for c in df.columns:
        norm = " ".join(str(c).strip().lower().split())
        if norm == target_norm:
            return c
    return None


def _clean_text(v):
    if pd.isna(v):
        return None
    s = str(v).strip()
    return s if s else None


def _clean_phone(v):
    if pd.isna(v):
        return None
    try:
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        if isinstance(v, (int,)):
            return str(v)
    except Exception:
        pass
    return str(v).strip().replace(".0", "")


def _clean_date(v):
    if pd.isna(v):
        return None
    try:
        return pd.to_datetime(v).date().isoformat()
    except Exception:
        return _clean_text(v)


def _clean_timestamp(v):
    if pd.isna(v):
        return None
    try:
        return pd.to_datetime(v).isoformat(sep=" ")
    except Exception:
        return _clean_text(v)


def import_responses(file_obj: str | Path | BinaryIO | bytes) -> dict:
    init_db()
    if isinstance(file_obj, bytes):
        source = BytesIO(file_obj)
    else:
        source = file_obj

    df = pd.read_excel(source, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]

    roll_col = _find_col(df, "Roll Number")
    name_col = _find_col(df, "Full Name")
    timestamp_col = _find_col(df, "Timestamp")
    if not roll_col or not name_col:
        raise ValueError("The workbook must contain 'Full Name' and 'Roll Number' columns.")

    # Keep latest response for duplicate roll/URN entries.
    if timestamp_col:
        df["__ts_sort"] = pd.to_datetime(df[timestamp_col], errors="coerce")
        df = df.sort_values("__ts_sort")

    conn = get_conn()
    roster = conn.execute("SELECT roll_no, urn FROM students").fetchall()
    roll_set = {int(r["roll_no"]) for r in roster}
    urn_map = {str(r["urn"]): int(r["roll_no"]) for r in roster}

    mapped = {}
    unmatched = []
    duplicates = set()

    for _, r in df.iterrows():
        raw = r[roll_col]
        key = None
        if not pd.isna(raw):
            try:
                key = str(int(float(raw)))
            except Exception:
                key = str(raw).strip()
        roll_no = None
        if key and key.isdigit() and int(key) in roll_set:
            roll_no = int(key)
        elif key in urn_map:
            roll_no = urn_map[key]
        if roll_no is None:
            unmatched.append(_clean_text(r[name_col]) or "Unknown")
            continue
        if roll_no in mapped:
            duplicates.add(roll_no)
        mapped[roll_no] = r

    cols = {
        "email": _find_col(df, "Email Address"),
        "dob": _find_col(df, "Date of birth"),
        "gender": _find_col(df, "Gender"),
        "blood_group": _find_col(df, "Blood Group"),
        "cpi": _find_col(df, "CPI"),
        "contact_no": _find_col(df, "Personal Mobile Number"),
        "father_name": _find_col(df, "Fathers Full Name"),
        "residence_address": _find_col(df, "Residence address"),
        "father_occupation": _find_col(df, "Father Occupation"),
        "father_mobile": _find_col(df, "Father Mobile Number"),
        "emergency_contact": _find_col(df, "Emergency Contact Phone Number"),
        "goals": _find_col(df, "Short-Term and Long-Term Goals (e.g., career advancement, skill building, or passing an exam)"),
        "interests_hobbies": _find_col(df, "Areas of Interest & Hobbies"),
        "strengths": _find_col(df, "Strengths"),
        "weaknesses": _find_col(df, "Weaknesses"),
        "medical_condition": _find_col(df, "Do you have any known medical conditions? If yes, please describe your medical condition"),
        "allergies": _find_col(df, "Do you have any food, drug, or environmental allergies?"),
        "emergency_medication": _find_col(df, "Do you carry any emergency medication? (e.g., Inhaler, EpiPen)"),
        "extra_info": _find_col(df, "Is there anything else you would like to share that was not covered in this form?"),
    }

    for roll_no, r in mapped.items():
        values = {
            "email": _clean_text(r[cols["email"]]) if cols["email"] else None,
            "dob": _clean_date(r[cols["dob"]]) if cols["dob"] else None,
            "gender": _clean_text(r[cols["gender"]]) if cols["gender"] else None,
            "blood_group": _clean_text(r[cols["blood_group"]]) if cols["blood_group"] else None,
            "cpi": None,
            "contact_no": _clean_phone(r[cols["contact_no"]]) if cols["contact_no"] else None,
            "father_name": _clean_text(r[cols["father_name"]]) if cols["father_name"] else None,
            "residence_address": _clean_text(r[cols["residence_address"]]) if cols["residence_address"] else None,
            "father_occupation": _clean_text(r[cols["father_occupation"]]) if cols["father_occupation"] else None,
            "father_mobile": _clean_phone(r[cols["father_mobile"]]) if cols["father_mobile"] else None,
            "emergency_contact": _clean_phone(r[cols["emergency_contact"]]) if cols["emergency_contact"] else None,
            "goals": _clean_text(r[cols["goals"]]) if cols["goals"] else None,
            "interests_hobbies": _clean_text(r[cols["interests_hobbies"]]) if cols["interests_hobbies"] else None,
            "strengths": _clean_text(r[cols["strengths"]]) if cols["strengths"] else None,
            "weaknesses": _clean_text(r[cols["weaknesses"]]) if cols["weaknesses"] else None,
            "medical_condition": _clean_text(r[cols["medical_condition"]]) if cols["medical_condition"] else None,
            "allergies": _clean_text(r[cols["allergies"]]) if cols["allergies"] else None,
            "emergency_medication": _clean_text(r[cols["emergency_medication"]]) if cols["emergency_medication"] else None,
            "extra_info": _clean_text(r[cols["extra_info"]]) if cols["extra_info"] else None,
            "response_timestamp": _clean_timestamp(r[timestamp_col]) if timestamp_col else datetime.now().isoformat(sep=" "),
        }
        if cols["cpi"] and not pd.isna(r[cols["cpi"]]):
            try:
                values["cpi"] = float(r[cols["cpi"]])
            except Exception:
                pass

        conn.execute(
            """
            UPDATE students SET
                email=?, cpi=?, dob=?, gender=?, blood_group=?, contact_no=COALESCE(?, contact_no),
                father_name=?, residence_address=?, father_occupation=?, father_mobile=?, emergency_contact=?,
                goals=?, interests_hobbies=?, strengths=?, weaknesses=?, medical_condition=?, allergies=?,
                emergency_medication=?, extra_info=?, response_timestamp=?, response_status='Received',
                updated_at=CURRENT_TIMESTAMP
            WHERE roll_no=?
            """,
            (
                values["email"], values["cpi"], values["dob"], values["gender"], values["blood_group"],
                values["contact_no"], values["father_name"], values["residence_address"], values["father_occupation"],
                values["father_mobile"], values["emergency_contact"], values["goals"], values["interests_hobbies"],
                values["strengths"], values["weaknesses"], values["medical_condition"], values["allergies"],
                values["emergency_medication"], values["extra_info"], values["response_timestamp"], roll_no,
            ),
        )

    # Reset status to Pending only for roster members that have never supplied a response.
    conn.execute("UPDATE students SET response_status='Pending' WHERE response_timestamp IS NULL")
    conn.commit()
    pending = [r["roll_no"] for r in conn.execute("SELECT roll_no FROM students WHERE response_status='Pending' ORDER BY roll_no")]
    conn.close()

    return {
        "mapped": len(mapped),
        "pending": pending,
        "duplicates": sorted(duplicates),
        "unmatched": unmatched,
    }
