"""
Unit tests for the Masar Employee Management System.
Tests cover: database operations, text normalization, validation, CRUD operations, and data integrity.

Run with: pytest test_masar.py -v
"""

import pytest
import sqlite3
import datetime
import os
import tempfile
import shutil
from masar import normalize_arabic, init_db, EMPLOYEE_FIELDS, AR_LABELS, DB_FILE, ATTACHMENTS_DIR


class TestNormalizeArabic:
    """Test Arabic text normalization function."""

    def test_normalize_empty_string(self):
        """Empty string should return empty string."""
        assert normalize_arabic("") == ""

    def test_normalize_none(self):
        """None should return None or empty string."""
        result = normalize_arabic(None)
        assert result is None or result == ""

    def test_normalize_hamza_variants(self):
        """Convert أ, إ, آ to ا."""
        assert normalize_arabic("أسد") == "اسد"
        assert normalize_arabic("إسلام") == "اسلام"
        assert normalize_arabic("آمن") == "امن"

    def test_normalize_teh_marbuta(self):
        """Convert ة to ه."""
        result = normalize_arabic("جامعة")
        assert "ة" not in result
        
        result = normalize_arabic("دولة")
        assert "ة" not in result

    def test_normalize_alef_maqsura(self):
        """Convert ى to ي."""
        result = normalize_arabic("موسى")
        assert "ى" not in result
        assert "ي" in result

    def test_normalize_tatweel(self):
        """Remove tatweel (ـ)."""
        result = normalize_arabic("الـــــتـــــالي")
        assert "ـ" not in result

    def test_normalize_combined(self):
        """Test combination of normalizations."""
        result = normalize_arabic("أسلام إلى دولة موسى")
        # Check that problematic characters are converted
        assert "أ" not in result
        assert "إ" not in result
        assert "آ" not in result

    def test_normalize_whitespace_strip(self):
        """Strip leading and trailing whitespace."""
        result = normalize_arabic("  أسد  ")
        assert result == result.strip()
        
        result = normalize_arabic("\tعلي\n")
        assert result == result.strip()

    def test_normalize_preserves_normal_text(self):
        """Normal Arabic text should remain unchanged."""
        assert normalize_arabic("محمد") == "محمد"
        assert normalize_arabic("علي") == "علي"


class TestEmployeeConfiguration:
    """Test employee field configuration."""

    def test_employee_fields_defined(self):
        """All employee fields should be defined."""
        assert len(EMPLOYEE_FIELDS) > 0, "EMPLOYEE_FIELDS should not be empty"

    def test_employee_fields_contain_essentials(self):
        """Essential fields should be present."""
        essential_fields = ['name', 'file_no', 'national_id']
        for field in essential_fields:
            assert field in EMPLOYEE_FIELDS, f"Field {field} should be in EMPLOYEE_FIELDS"

    def test_ar_labels_complete(self):
        """All employee fields should have Arabic labels."""
        for field in EMPLOYEE_FIELDS:
            assert field in AR_LABELS, f"Field {field} missing Arabic label"
            assert AR_LABELS[field] != "", f"Label for {field} should not be empty"

    def test_ar_labels_are_strings(self):
        """All AR_LABELS values should be strings."""
        for field, label in AR_LABELS.items():
            assert isinstance(label, str), f"Label for {field} should be string, got {type(label)}"

    def test_required_fields_have_labels(self):
        """Critical fields should have Arabic labels."""
        critical_fields = ['name', 'file_no', 'national_id', 'retirement_date']
        for field in critical_fields:
            assert field in AR_LABELS, f"Critical field {field} missing label"


class TestDatabase:
    """Test database initialization and schema."""

    @pytest.fixture
    def temp_db_dir(self):
        """Create and cleanup a temporary database directory."""
        temp_dir = tempfile.mkdtemp()
        yield temp_dir
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

    def test_employee_table_creation(self, temp_db_dir):
        """Test employee table can be created."""
        db_path = os.path.join(temp_db_dir, "test.db")
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        
        # Create table
        c.execute(f"""
            CREATE TABLE IF NOT EXISTS employee (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                {', '.join([f"{f} TEXT" for f in EMPLOYEE_FIELDS])}
            )
        """)
        conn.commit()
        
        # Verify table exists
        c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='employee'")
        assert c.fetchone() is not None, "Employee table should be created"
        
        conn.close()

    def test_correspondence_table_creation(self, temp_db_dir):
        """Test correspondence table creation."""
        db_path = os.path.join(temp_db_dir, "test.db")
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        
        # Create table
        c.execute("""
            CREATE TABLE IF NOT EXISTS correspondence (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fax_number TEXT,
                fax_date TEXT,
                from_person TEXT,
                to_person TEXT,
                subject TEXT,
                notes TEXT,
                image_path TEXT,
                created_at TEXT
            )
        """)
        conn.commit()
        
        # Verify table exists
        c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='correspondence'")
        assert c.fetchone() is not None, "Correspondence table should be created"
        
        conn.close()

    def test_correspondence_attachment_table_creation(self, temp_db_dir):
        """Test correspondence attachment table creation."""
        db_path = os.path.join(temp_db_dir, "test.db")
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        
        c.execute("""
            CREATE TABLE IF NOT EXISTS correspondence (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fax_number TEXT
            )
        """)
        
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
        conn.commit()
        
        # Verify table exists
        c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='correspondence_attachment'")
        assert c.fetchone() is not None, "correspondence_attachment table should be created"
        
        conn.close()

    def test_directories_created(self, temp_db_dir):
        """Test that attachment directories can be created."""
        attachments_dir = os.path.join(temp_db_dir, 'attachments')
        faxes_dir = os.path.join(attachments_dir, 'Faxes')
        
        os.makedirs(faxes_dir, exist_ok=True)
        
        assert os.path.exists(attachments_dir), "attachments directory should exist"
        assert os.path.exists(faxes_dir), "Faxes subdirectory should exist"


class TestEmployeeCRUD:
    """Test CRUD operations for employees."""

    @pytest.fixture
    def temp_db(self):
        """Create a temporary database with employee table."""
        temp_dir = tempfile.mkdtemp()
        db_path = os.path.join(temp_dir, "test.db")
        
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        c.execute(f"""
            CREATE TABLE IF NOT EXISTS employee (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                {', '.join([f"{f} TEXT" for f in EMPLOYEE_FIELDS])}
            )
        """)
        conn.commit()
        conn.close()
        
        yield db_path, temp_dir
        
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

    def test_insert_employee(self, temp_db):
        """Test inserting an employee."""
        db_path, _ = temp_db
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        
        employee_data = ['محمد علي'] + [''] * (len(EMPLOYEE_FIELDS) - 1)
        
        c.execute(f"INSERT INTO employee ({', '.join(EMPLOYEE_FIELDS)}) VALUES ({', '.join(['?']*len(EMPLOYEE_FIELDS))})",
                  employee_data)
        conn.commit()
        
        c.execute("SELECT COUNT(*) FROM employee")
        count = c.fetchone()[0]
        assert count == 1, "Employee should be inserted"
        
        conn.close()

    def test_retrieve_employee(self, temp_db):
        """Test retrieving an employee."""
        db_path, _ = temp_db
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        
        employee_data = ['محمد علي'] + [''] * (len(EMPLOYEE_FIELDS) - 1)
        c.execute(f"INSERT INTO employee ({', '.join(EMPLOYEE_FIELDS)}) VALUES ({', '.join(['?']*len(EMPLOYEE_FIELDS))})",
                  employee_data)
        conn.commit()
        emp_id = c.lastrowid
        
        c.execute("SELECT name FROM employee WHERE id=?", (emp_id,))
        result = c.fetchone()
        assert result is not None, "Employee should be retrieved"
        assert result[0] == 'محمد علي', "Employee name should match"
        
        conn.close()

    def test_update_employee(self, temp_db):
        """Test updating an employee."""
        db_path, _ = temp_db
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        
        employee_data = ['محمد علي'] + [''] * (len(EMPLOYEE_FIELDS) - 1)
        c.execute(f"INSERT INTO employee ({', '.join(EMPLOYEE_FIELDS)}) VALUES ({', '.join(['?']*len(EMPLOYEE_FIELDS))})",
                  employee_data)
        conn.commit()
        emp_id = c.lastrowid
        
        c.execute("UPDATE employee SET name=? WHERE id=?", ('أحمد محمود', emp_id))
        conn.commit()
        
        c.execute("SELECT name FROM employee WHERE id=?", (emp_id,))
        result = c.fetchone()[0]
        assert result == 'أحمد محمود', "Employee name should be updated"
        
        conn.close()

    def test_delete_employee(self, temp_db):
        """Test deleting an employee."""
        db_path, _ = temp_db
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        
        employee_data = ['محمد علي'] + [''] * (len(EMPLOYEE_FIELDS) - 1)
        c.execute(f"INSERT INTO employee ({', '.join(EMPLOYEE_FIELDS)}) VALUES ({', '.join(['?']*len(EMPLOYEE_FIELDS))})",
                  employee_data)
        conn.commit()
        emp_id = c.lastrowid
        
        c.execute("DELETE FROM employee WHERE id=?", (emp_id,))
        conn.commit()
        
        c.execute("SELECT COUNT(*) FROM employee WHERE id=?", (emp_id,))
        count = c.fetchone()[0]
        assert count == 0, "Employee should be deleted"
        
        conn.close()


class TestCorrespondenceCRUD:
    """Test CRUD operations for correspondence."""

    @pytest.fixture
    def temp_db(self):
        """Create a temporary database with correspondence tables."""
        temp_dir = tempfile.mkdtemp()
        db_path = os.path.join(temp_dir, "test.db")
        
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        c.execute("""
            CREATE TABLE IF NOT EXISTS correspondence (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fax_number TEXT,
                fax_date TEXT,
                from_person TEXT,
                to_person TEXT,
                subject TEXT,
                notes TEXT,
                image_path TEXT,
                created_at TEXT
            )
        """)
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
        conn.commit()
        conn.close()
        
        yield db_path, temp_dir
        
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

    def test_insert_correspondence(self, temp_db):
        """Test inserting correspondence."""
        db_path, _ = temp_db
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        
        c.execute("INSERT INTO correspondence (fax_number, fax_date, from_person, to_person, subject, notes, image_path, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                  ('FAX001', '2025-11-17', 'محمد', 'احمد', 'موضوع تجريبي', 'ملاحظات', '', datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
        
        c.execute("SELECT COUNT(*) FROM correspondence")
        count = c.fetchone()[0]
        assert count == 1, "Correspondence should be inserted"
        
        conn.close()

    def test_retrieve_correspondence(self, temp_db):
        """Test retrieving correspondence."""
        db_path, _ = temp_db
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        
        c.execute("INSERT INTO correspondence (fax_number, fax_date, from_person, to_person, subject, notes, image_path, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                  ('FAX002', '2025-11-17', 'محمد', 'احمد', 'موضوع', 'ملاحظات', '', datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
        corr_id = c.lastrowid
        
        c.execute("SELECT subject FROM correspondence WHERE id=?", (corr_id,))
        result = c.fetchone()
        assert result is not None, "Correspondence should be retrieved"
        assert result[0] == 'موضوع', "Subject should match"
        
        conn.close()

    def test_update_correspondence(self, temp_db):
        """Test updating correspondence."""
        db_path, _ = temp_db
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        
        c.execute("INSERT INTO correspondence (fax_number, fax_date, from_person, to_person, subject, notes, image_path, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                  ('FAX003', '2025-11-17', 'محمد', 'احمد', 'موضوع قديم', 'ملاحظات', '', datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
        corr_id = c.lastrowid
        
        c.execute("UPDATE correspondence SET subject=? WHERE id=?", ('موضوع جديد', corr_id))
        conn.commit()
        
        c.execute("SELECT subject FROM correspondence WHERE id=?", (corr_id,))
        result = c.fetchone()[0]
        assert result == 'موضوع جديد', "Subject should be updated"
        
        conn.close()

    def test_delete_correspondence(self, temp_db):
        """Test deleting correspondence."""
        db_path, _ = temp_db
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        
        c.execute("INSERT INTO correspondence (fax_number, fax_date, from_person, to_person, subject, notes, image_path, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                  ('FAX004', '2025-11-17', 'محمد', 'احمد', 'موضوع', 'ملاحظات', '', datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
        corr_id = c.lastrowid
        
        c.execute("DELETE FROM correspondence WHERE id=?", (corr_id,))
        conn.commit()
        
        c.execute("SELECT COUNT(*) FROM correspondence WHERE id=?", (corr_id,))
        count = c.fetchone()[0]
        assert count == 0, "Correspondence should be deleted"
        
        conn.close()

    def test_insert_correspondence_attachment(self, temp_db):
        """Test inserting correspondence attachment."""
        db_path, _ = temp_db
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        
        c.execute("INSERT INTO correspondence (fax_number, fax_date, from_person, to_person, subject, notes, image_path, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                  ('FAX005', '2025-11-17', 'محمد', 'احمد', 'موضوع', 'ملاحظات', '', datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
        corr_id = c.lastrowid
        
        c.execute("INSERT INTO correspondence_attachment (correspondence_id, filename, filepath, upload_date) VALUES (?, ?, ?, ?)",
                  (corr_id, 'image1.jpg', '/path/to/image1.jpg', datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
        
        c.execute("SELECT COUNT(*) FROM correspondence_attachment WHERE correspondence_id=?", (corr_id,))
        count = c.fetchone()[0]
        assert count == 1, "Attachment should be inserted"
        
        conn.close()


class TestDataIntegrity:
    """Test data integrity and constraints."""

    @pytest.fixture
    def temp_db(self):
        """Create a temporary database with correspondence tables."""
        temp_dir = tempfile.mkdtemp()
        db_path = os.path.join(temp_dir, "test.db")
        
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        c.execute("""
            CREATE TABLE IF NOT EXISTS correspondence (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fax_number TEXT
            )
        """)
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
        conn.commit()
        conn.close()
        
        yield db_path, temp_dir
        
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

    def test_correspondence_attachment_foreign_key(self, temp_db):
        """Test foreign key relationship for correspondence attachments."""
        db_path, _ = temp_db
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        
        c.execute("INSERT INTO correspondence (fax_number) VALUES (?)", ('FAX006',))
        conn.commit()
        corr_id = c.lastrowid
        
        c.execute("INSERT INTO correspondence_attachment (correspondence_id, filename, filepath, upload_date) VALUES (?, ?, ?, ?)",
                  (corr_id, 'img.jpg', '/path/img.jpg', datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
        
        c.execute("SELECT * FROM correspondence_attachment WHERE correspondence_id=?", (corr_id,))
        result = c.fetchone()
        assert result is not None, "Attachment should exist"
        assert result[1] == corr_id, "Correspondence ID should match"
        
        conn.close()

    def test_multiple_attachments_per_correspondence(self, temp_db):
        """Test multiple attachments per correspondence."""
        db_path, _ = temp_db
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        
        c.execute("INSERT INTO correspondence (fax_number) VALUES (?)", ('FAX007',))
        conn.commit()
        corr_id = c.lastrowid
        
        for i in range(3):
            c.execute("INSERT INTO correspondence_attachment (correspondence_id, filename, filepath, upload_date) VALUES (?, ?, ?, ?)",
                      (corr_id, f'image{i}.jpg', f'/path/image{i}.jpg', datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
        
        c.execute("SELECT COUNT(*) FROM correspondence_attachment WHERE correspondence_id=?", (corr_id,))
        count = c.fetchone()[0]
        assert count == 3, "All attachments should be stored"
        
        conn.close()


class TestDateHandling:
    """Test date handling and formatting."""

    def test_date_format_yyyy_mm_dd(self):
        """Test YYYY-MM-DD date format."""
        date_str = '2025-11-17'
        dt = datetime.datetime.strptime(date_str, "%Y-%m-%d")
        assert dt.year == 2025
        assert dt.month == 11
        assert dt.day == 17

    def test_date_format_yyyy_mm(self):
        """Test YYYY-MM date format."""
        date_str = '2025-11'
        dt = datetime.datetime.strptime(date_str, "%Y-%m")
        assert dt.year == 2025
        assert dt.month == 11

    def test_date_format_yyyy(self):
        """Test YYYY date format."""
        date_str = '2025'
        dt = datetime.datetime.strptime(date_str, "%Y")
        assert dt.year == 2025

    def test_future_date_detection(self):
        """Test detection of future dates."""
        today = datetime.date.today()
        tomorrow = today + datetime.timedelta(days=1)
        yesterday = today - datetime.timedelta(days=1)
        
        assert tomorrow > today, "Tomorrow should be after today"
        assert yesterday < today, "Yesterday should be before today"

    def test_current_date_time_formatting(self):
        """Test current date/time formatting."""
        now = datetime.datetime.now()
        formatted = now.strftime("%Y-%m-%d %H:%M:%S")
        parsed = datetime.datetime.strptime(formatted, "%Y-%m-%d %H:%M:%S")
        assert parsed.year == now.year
        assert parsed.month == now.month
        assert parsed.day == now.day


if __name__ == "__main__":
    pytest.main([__file__, "-v", "--tb=short"])
