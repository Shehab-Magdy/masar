# database/db_manager.py
import sqlite3
import os
from utils.constants import DB_FILE, ATTACHMENTS_DIR, EMPLOYEE_FIELDS

def init_db():
    """
    Initializes the database by creating the necessary tables if they don't exist.
    Now includes filetype and upload_date columns in the attachment table.
    """
    try:
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        # Add insurance_doc and retirement_date to the table creation
        c.execute(f"""
            CREATE TABLE IF NOT EXISTS employee (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                {', '.join([f"{f} TEXT" for f in EMPLOYEE_FIELDS])}
            )
        """)
        # Try to add the columns if missing (for upgrades)
        try:
            c.execute("ALTER TABLE employee ADD COLUMN retirement_date TEXT")
        except sqlite3.OperationalError:
            pass
        try:
            c.execute("ALTER TABLE employee ADD COLUMN insurance_doc TEXT")
        except sqlite3.OperationalError:
            pass
        try:
            c.execute("ALTER TABLE employee ADD COLUMN relative_name TEXT")
        except sqlite3.OperationalError:
            pass
        try:
            c.execute("ALTER TABLE employee ADD COLUMN relative_phone TEXT")
        except sqlite3.OperationalError:
            pass
        try:
            c.execute("ALTER TABLE employee ADD COLUMN relative_relation TEXT")
        except sqlite3.OperationalError:
            pass
        c.execute("""
            CREATE TABLE IF NOT EXISTS attachment (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id INTEGER,
                filename TEXT,
                filepath TEXT,
                filetype TEXT,
                upload_date TEXT,
                is_photo INTEGER DEFAULT 0,
                FOREIGN KEY(employee_id) REFERENCES employee(id)
            )
        """)
        # Try to add columns if not exist (for upgrades)
        try:
            c.execute("ALTER TABLE attachment ADD COLUMN filetype TEXT")
        except sqlite3.OperationalError:
            pass
        try:
            c.execute("ALTER TABLE attachment ADD COLUMN upload_date TEXT")
        except sqlite3.OperationalError:
            pass
        # Create correspondence table for faxes/letters
        try:
            c.execute("""
                CREATE TABLE IF NOT EXISTS correspondence (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    fax_number TEXT,
                    correspondence_type TEXT NOT NULL DEFAULT 'صادر',
                    fax_date TEXT,
                    from_person TEXT,
                    to_person TEXT,
                    subject TEXT,
                    notes TEXT,
                    image_path TEXT,
                    created_at TEXT
                )
            """)
        except Exception:
            pass
        try:
            c.execute("ALTER TABLE correspondence ADD COLUMN correspondence_type TEXT NOT NULL DEFAULT 'صادر'")
            # Set all existing rows to 'صادر'
            c.execute("UPDATE correspondence SET correspondence_type = 'صادر' WHERE correspondence_type IS NULL OR correspondence_type = ''")
        except sqlite3.OperationalError:
            pass
        # Attachments for correspondence (multiple images per correspondence)
        try:
            c.execute("""
                CREATE TABLE IF NOT EXISTS correspondence_attachment (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    correspondence_id INTEGER,
                    filename TEXT,
                    filepath TEXT,
                    upload_date TEXT,
                    FOREIGN KEY(correspondence_id) REFERENCES correspondence(id)
                )
            """)
        except Exception:
            pass
        conn.commit()
        conn.close()
        if not os.path.exists(ATTACHMENTS_DIR):
            os.makedirs(ATTACHMENTS_DIR)
        # ensure faxes attachments folder exists
        faxes_dir = os.path.join(ATTACHMENTS_DIR, 'Faxes')
        if not os.path.exists(faxes_dir):
            os.makedirs(faxes_dir)
    except Exception as e:
        print("Database initialization error:", e)
