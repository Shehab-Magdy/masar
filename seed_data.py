#!/usr/bin/env python3
"""
Development/testing helper: fills the Masar database with sample employees
(covering every retirement-date scenario the dashboard/reports distinguish
between) and sample correspondence/fax records spread across different dates.

Run from the project root (same place you'd run `python main.py`), since it
uses the same masar.db / attachments/ paths as the app itself.

Usage:
    python seed_data.py                        # 60 employees, 60 faxes
    python seed_data.py --employees 100 --faxes 80
    python seed_data.py --reset                # wipe employee/attachment/correspondence data first, then seed
"""
import argparse
import datetime
import random
import sqlite3

from database.db_manager import init_db
from utils.constants import DB_FILE, EMPLOYEE_FIELDS

DEPARTMENTS = [
    "الإدارة", "الموارد البشرية", "المالية", "الصيانة",
    "الإنتاج", "الجودة", "المشتريات", "أمن الورش",
]
JOB_TITLES = ["فني", "مهندس", "مراقب جودة", "محاسب", "سائق", "عامل", "مشرف", "كاتب"]
QUALIFICATIONS = ["دبلوم صناعي", "بكالوريوس", "ثانوية عامة", "دبلوم فني", "ماجستير"]
GRADES = ["الأولى", "الثانية", "الثالثة", "الرابعة"]
FUNCTIONAL_GROUPS = ["فنية", "إدارية", "خدمية"]
TYPE_GROUPS = ["عمال", "فنيين", "إداريين"]
RELATIONS = ["الزوج/ة", "الأب", "الأم", "الابن/ة", "الأخ/ت"]
FIRST_NAMES = [
    "محمد", "أحمد", "محمود", "علي", "حسن", "إبراهيم", "خالد", "عمر",
    "فاطمة", "منى", "سارة", "هدى", "ياسمين", "نور",
]
LAST_NAMES = [
    "السيد", "عبد الله", "حسين", "إبراهيم", "الشريف", "عطية",
    "راشد", "مصطفى", "فهمي", "عبد الرازق",
]
FROM_TO_NAMES = [
    "محمد علي", "أحمد حسن", "إدارة الورشة", "قسم المشتريات",
    "مكتب المدير", "الشؤون القانونية",
]
CORR_SUBJECTS = [
    "طلب صيانة", "تقرير دوري", "إشعار اجتماع", "طلب توريد مستلزمات",
    "استفسار بخصوص", "رد على طلب", "تعميم إداري", "إحالة ملف",
]


def random_date(start, end):
    """A random date between start and end (inclusive), clamped if end < start."""
    delta_days = (end - start).days
    return start + datetime.timedelta(days=random.randint(0, max(delta_days, 0)))


def retirement_date_scenarios(today):
    """
    Cycles through every retirement_date scenario the app's dashboard/reports
    distinguish between: none set, already retired, retiring later this year,
    retiring within 3/6 months, retiring exactly today (boundary case), and
    retiring years from now.
    """
    scenarios = [
        lambda: "",
        lambda: (today - datetime.timedelta(days=random.randint(365, 365 * 10))).strftime("%Y-%m-%d"),
        lambda: random_date(today, today.replace(month=12, day=31)).strftime("%Y-%m-%d"),
        lambda: random_date(today, today + datetime.timedelta(days=90)).strftime("%Y-%m-%d"),
        lambda: random_date(today, today + datetime.timedelta(days=180)).strftime("%Y-%m-%d"),
        lambda: today.strftime("%Y-%m-%d"),
        lambda: random_date(
            today + datetime.timedelta(days=365 * 2), today + datetime.timedelta(days=365 * 5)
        ).strftime("%Y-%m-%d"),
    ]
    i = 0
    while True:
        yield scenarios[i % len(scenarios)]()
        i += 1


def seed_employees(conn, count, start_file_no, start_national_id):
    today = datetime.date.today()
    scenarios = retirement_date_scenarios(today)
    rows = []
    for i in range(count):
        file_no = str(start_file_no + i)
        national_id = str(start_national_id + i).zfill(14)
        name = f"{random.choice(FIRST_NAMES)} {random.choice(LAST_NAMES)}"
        hire_date = random_date(datetime.date(1990, 1, 1), today - datetime.timedelta(days=365)).strftime("%Y-%m-%d")
        birth_date = random_date(datetime.date(1960, 1, 1), today - datetime.timedelta(days=365 * 20)).strftime("%Y-%m-%d")
        grade_date = random_date(datetime.date(2000, 1, 1), today).strftime("%Y-%m-%d")
        retirement_date = next(scenarios)
        # ~30% of seeded employees are missing their insurance doc, to exercise
        # the dashboard's "missing-document" stat.
        insurance_doc = "" if random.random() < 0.3 else f"DOC-{file_no}"
        values = {
            "name": name,
            "grade": random.choice(GRADES),
            "grade_date": grade_date,
            "hire_date": hire_date,
            "file_no": file_no,
            "qualification": random.choice(QUALIFICATIONS),
            "functional_group": random.choice(FUNCTIONAL_GROUPS),
            "type_group": random.choice(TYPE_GROUPS),
            "job_title": random.choice(JOB_TITLES),
            "department": random.choice(DEPARTMENTS),
            "current_work": "عمل حالي تجريبي",
            "birth_date": birth_date,
            "retirement_date": retirement_date,
            "insurance_no": str(random.randint(10_000_000, 99_999_999)),
            "national_id": national_id,
            "address": "عنوان تجريبي - القاهرة",
            "phone": "01" + str(random.randint(0, 9)) + str(random.randint(10_000_000, 99_999_999)),
            "insurance_doc": insurance_doc,
            "notes": "بيانات تجريبية (seed_data.py)",
            "relative_name": f"{random.choice(FIRST_NAMES)} {random.choice(LAST_NAMES)}",
            "relative_phone": "01" + str(random.randint(0, 9)) + str(random.randint(10_000_000, 99_999_999)),
            "relative_relation": random.choice(RELATIONS),
        }
        rows.append(tuple(values[f] for f in EMPLOYEE_FIELDS))

    c = conn.cursor()
    c.executemany(
        f"INSERT INTO employee ({', '.join(EMPLOYEE_FIELDS)}) VALUES ({', '.join(['?'] * len(EMPLOYEE_FIELDS))})",
        rows,
    )
    conn.commit()
    return len(rows)


def seed_faxes(conn, count, start_fax_no):
    today = datetime.date.today()
    rows = []
    for i in range(count):
        fax_number = str(start_fax_no + i)
        corr_type = random.choice(["صادر", "وارد"])
        fax_date = random_date(today - datetime.timedelta(days=365 * 2), today).strftime("%Y-%m-%d")
        from_person = random.choice(FROM_TO_NAMES)
        to_person = random.choice(FROM_TO_NAMES)
        subject = f"{random.choice(CORR_SUBJECTS)} رقم {i + 1}"
        notes = "بيانات تجريبية (seed_data.py)"
        created_at = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        rows.append((fax_number, corr_type, fax_date, from_person, to_person, subject, notes, "", created_at))

    c = conn.cursor()
    c.executemany(
        "INSERT INTO correspondence "
        "(fax_number, correspondence_type, fax_date, from_person, to_person, subject, notes, image_path, created_at) "
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
        rows,
    )
    conn.commit()
    return len(rows)


def reset_tables(conn):
    c = conn.cursor()
    for table in ("attachment", "correspondence_attachment", "correspondence", "employee"):
        c.execute(f"DELETE FROM {table}")
    conn.commit()


def main():
    parser = argparse.ArgumentParser(
        description="Seed the Masar database with sample employees and correspondence for testing."
    )
    parser.add_argument("--employees", type=int, default=60, help="Number of employees to insert (default: 60)")
    parser.add_argument("--faxes", type=int, default=60, help="Number of correspondence/fax records to insert (default: 60)")
    parser.add_argument("--reset", action="store_true", help="Delete existing employee/attachment/correspondence data first")
    args = parser.parse_args()

    init_db()
    conn = sqlite3.connect(DB_FILE)

    if args.reset:
        reset_tables(conn)
        print("Existing employee/correspondence data cleared.")

    # Start numbering above anything already in the DB, so seeded file_no /
    # national_id / fax_number never collide with real or previously-seeded rows.
    c = conn.cursor()
    c.execute("SELECT COALESCE(MAX(CAST(file_no AS INTEGER)), 9000) FROM employee WHERE file_no GLOB '[0-9]*'")
    start_file_no = (c.fetchone()[0] or 9000) + 1
    c.execute("SELECT COALESCE(MAX(CAST(fax_number AS INTEGER)), 9000) FROM correspondence WHERE fax_number GLOB '[0-9]*'")
    start_fax_no = (c.fetchone()[0] or 9000) + 1
    start_national_id = 10_000_000_000_000 + start_file_no  # 14 digits, offset to stay unique

    n_emp = seed_employees(conn, args.employees, start_file_no, start_national_id)
    n_fax = seed_faxes(conn, args.faxes, start_fax_no)

    conn.close()
    print(f"Inserted {n_emp} employees and {n_fax} correspondence records into {DB_FILE}.")


if __name__ == "__main__":
    main()
