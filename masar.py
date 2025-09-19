import sys
import os
import sqlite3
import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget, QVBoxLayout, QLabel, QPushButton,
    QTableWidget, QTableWidgetItem, QLineEdit, QHBoxLayout, QFileDialog, QListWidget,
    QMessageBox, QTextEdit, QFormLayout
)
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt
from weasyprint import HTML, CSS
import mimetypes
import re

DB_FILE = "masar.db"
ATTACHMENTS_DIR = "attachments"

AR_LABELS = {
    "name": "الاسم",
    "grade": "الدرجة",
    "grade_date": "تاريخ الحصول عليها",
    "hire_date": "تاريخ التعيين",
    "file_no": "رقم الملف",
    "qualification": "المؤهل",
    "functional_group": "مجموعة وظيفية",
    "type_group": "مجموعة نوعية",
    "job_title": "المسمى الوظيفي",
    "department": "القسم",
    "current_work": "العمل القائم به",
    "birth_date": "تاريخ الميلاد",
    "insurance_no": "رقم تأميني",
    "national_id": "رقم قومي",
    "address": "عنوان حالي",
    "phone": "رقم التليفون",
    "notes": "ملاحظات",
    "attachments": "ملفات مرتبطة",
    "personal_photo": "صورة شخصية"
}

EMPLOYEE_FIELDS = [
    "name", "grade", "grade_date", "hire_date", "file_no", "qualification",
    "functional_group", "type_group", "job_title", "department", "current_work",
    "birth_date", "insurance_no", "national_id", "address", "phone", "notes"
]

def init_db():
    """
    Initializes the database by creating the necessary tables if they don't exist.
    Now includes filetype and upload_date columns in the attachment table.
    """
    try:
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute(f"""
            CREATE TABLE IF NOT EXISTS employee (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                {', '.join([f"{f} TEXT" for f in EMPLOYEE_FIELDS])}
            )
        """)
        # Add new columns if not exist (for upgrades)
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
        # Try to add columns if missing (safe for existing DBs)
        try:
            c.execute("ALTER TABLE attachment ADD COLUMN filetype TEXT")
        except sqlite3.OperationalError:
            pass
        try:
            c.execute("ALTER TABLE attachment ADD COLUMN upload_date TEXT")
        except sqlite3.OperationalError:
            pass
        conn.commit()
        conn.close()
        if not os.path.exists(ATTACHMENTS_DIR):
            os.makedirs(ATTACHMENTS_DIR)
    except Exception as e:
        print("Database initialization error:", e)

def normalize_arabic(text: str) -> str:
    """
    Normalize Arabic text before saving to DB or searching.
    Converts أ, إ, آ to ا
    Converts ة to ه (optional, depends on your needs)
    Converts ى to ي
    Removes tatweel (ـ)
    """
    if not text:
        return text

    replacements = {
        "أ": "ا",
        "إ": "ا",
        "آ": "ا",
        "ة": "ه",  # optional, if you want unify it
        "ى": "ي",
    }

    # Remove tatweel (ـ)
    text = text.replace("ـ", "")

    for src, target in replacements.items():
        text = text.replace(src, target)

    return text.strip()

def get_employee_folder(file_no):
    """
    Returns the path to the employee's attachment folder based on file_no.
    """
    folder = os.path.join(ATTACHMENTS_DIR, str(file_no))
    if not os.path.exists(folder):
        os.makedirs(folder)
    return folder

class MasarMainWindow(QMainWindow):
    def __init__(self):
        """
        Initializes the main window of the application.

        Sets the window title, geometry, and icon.
        Connects to the database and sets up the tab widget with the three main tabs:
        "لوحة التحكم", "الموظفين", and "التقارير".
        """
        super().__init__()
        self.setWindowTitle("مسار - إدارة الموظفين")
        self.setGeometry(100, 100, 1100, 700)
        # Set window icon
        from PyQt5.QtGui import QIcon
        self.setWindowIcon(QIcon(os.path.join(os.path.dirname(__file__), "masar.ico")))
        self.conn = sqlite3.connect(DB_FILE)
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)
        self.tabs.addTab(DashboardTab(self.conn), "لوحة التحكم")
        self.tabs.addTab(EmployeeTab(self.conn), "الموظفين")
        self.tabs.addTab(ReportTab(self.conn), "التقارير")

class DashboardTab(QWidget):
    def __init__(self, conn):
        """
        Initializes the dashboard tab.

        Sets the window title and layout. It also refreshes the counts of employees, departments, and attachments.

        :param conn: The database connection.
        :type conn: sqlite3.Connection
        """
        super().__init__()
        self.conn = conn
        layout = QVBoxLayout()
        self.lbl_title = QLabel("لوحة التحكم")
        self.lbl_title.setStyleSheet("font-size:24px; font-weight:bold; color:#1976d2;")
        layout.addWidget(self.lbl_title)
        self.lbl_emp = QLabel()
        self.lbl_dept = QLabel()
        self.lbl_att = QLabel()
        layout.addWidget(self.lbl_emp)
        layout.addWidget(self.lbl_dept)
        layout.addWidget(self.lbl_att)
        self.btn_refresh = QPushButton("تحديث")
        self.btn_refresh.clicked.connect(self.refresh_counts)
        layout.addWidget(self.btn_refresh)
        self.setLayout(layout)
        self.refresh_counts()

    def refresh_counts(self):
        """
        Refreshes the counts of employees, departments, and attachments on the dashboard tab.

        Queries the database to get the counts and updates the labels accordingly.
        """
        c = self.conn.cursor()
        c.execute("SELECT COUNT(*) FROM employee")
        emp_count = c.fetchone()[0]
        c.execute("SELECT COUNT(DISTINCT department) FROM employee")
        dept_count = c.fetchone()[0]
        c.execute("SELECT COUNT(*) FROM attachment")
        att_count = c.fetchone()[0]
        self.lbl_emp.setText(f"عدد الموظفين: {emp_count}")
        self.lbl_dept.setText(f"عدد الأقسام: {dept_count}")
        self.lbl_att.setText(f"عدد الملفات المرتبطة: {att_count}")

class EmployeeTab(QWidget):
    def __init__(self, conn):
        """
        Initializes the employee tab.

        Sets up the UI layout and connects the signals of the UI elements to their respective slots.
        Loads the employee data from the database into the table widget.

        :param conn: The database connection.
        :type conn: sqlite3.Connection
        """
        super().__init__()
        self.conn = conn
        self.selected_emp_id = None
        self.attachments = []
        self.photo_path = None
        layout = QVBoxLayout()
        search_layout = QHBoxLayout()
        self.search_field = QLineEdit()
        self.search_field.setPlaceholderText("بحث بالاسم أو القسم أو الرقم القومى أو رقم الملف...")
        self.search_field.textChanged.connect(lambda: self.search_employees(self.search_field.text()))
        search_layout.addWidget(QLabel("بحث:"))
        search_layout.addWidget(self.search_field)
        layout.addLayout(search_layout)
        main_layout = QHBoxLayout()
        self.table = QTableWidget()
        self.table.setColumnCount(len(EMPLOYEE_FIELDS))
        self.table.setHorizontalHeaderLabels([AR_LABELS[f] for f in EMPLOYEE_FIELDS])
        self.table.setSelectionBehavior(self.table.SelectRows)
        self.table.cellClicked.connect(self.on_row_select)
        main_layout.addWidget(self.table)
        form_layout = QFormLayout()
        self.form_fields = {f: QLineEdit() for f in EMPLOYEE_FIELDS}
        for f in EMPLOYEE_FIELDS:
            form_layout.addRow(AR_LABELS[f], self.form_fields[f])
        self.attach_list = QListWidget()
        form_layout.addRow(AR_LABELS["attachments"], self.attach_list)
        self.btn_attach = QPushButton("رفع ملفات")
        self.btn_attach.clicked.connect(self.upload_files)
        form_layout.addRow(self.btn_attach)
        self.btn_photo = QPushButton("رفع صورة")
        self.btn_photo.clicked.connect(self.upload_photo)
        form_layout.addRow(AR_LABELS["personal_photo"], self.btn_photo)
        self.photo_label = QLabel()
        form_layout.addRow(self.photo_label)
        btns_layout = QHBoxLayout()
        self.btn_add = QPushButton("إضافة")
        self.btn_add.clicked.connect(self.add_employee)
        btns_layout.addWidget(self.btn_add)
        self.btn_edit = QPushButton("تعديل")
        self.btn_edit.clicked.connect(self.edit_employee)
        btns_layout.addWidget(self.btn_edit)
        self.btn_delete = QPushButton("حذف")
        self.btn_delete.clicked.connect(self.delete_employee)
        btns_layout.addWidget(self.btn_delete)
        self.btn_clear = QPushButton("مسح")
        self.btn_clear.clicked.connect(self.clear_form)
        btns_layout.addWidget(self.btn_clear)
        form_layout.addRow(btns_layout)
        form_widget = QWidget()
        form_widget.setLayout(form_layout)
        main_layout.addWidget(form_widget)
        layout.addLayout(main_layout)
        self.setLayout(layout)
        self.load_employees()

    def load_employees(self):
        """
        Loads all employees from the database into the table widget.

        Clears the table widget, then executes a SELECT query to retrieve all employee records.
        For each record, inserts a new row into the table widget and sets the values of the row
        according to the record's fields.

        :return: None
        :rtype: NoneType
        """
        self.table.setRowCount(0)
        c = self.conn.cursor()
        c.execute(f"SELECT id, {', '.join(EMPLOYEE_FIELDS)} FROM employee")
        for row in c.fetchall():
            row_idx = self.table.rowCount()
            self.table.insertRow(row_idx)
            for col_idx, val in enumerate(row[1:]):
                self.table.setItem(row_idx, col_idx, QTableWidgetItem(str(val)))
            self.table.setVerticalHeaderItem(row_idx, QTableWidgetItem(str(row[0])))

    def on_row_select(self, row, col):
        """
        Called when a row in the table widget is selected.

        Retrieves the employee ID from the selected row, then queries the database to retrieve
        the employee record. It then populates the form fields with the retrieved data and
        sets the selected employee ID attribute. Finally, it calls the load_attachments method to
        load the attachments for the selected employee.

        :param row: The row index of the selected row.
        :type row: int
        :param col: The column index of the selected row.
        :type col: int
        :return: None
        :rtype: NoneType
        """
        emp_id = self.table.verticalHeaderItem(row).text()
        c = self.conn.cursor()
        c.execute(f"SELECT {', '.join(EMPLOYEE_FIELDS)} FROM employee WHERE id=?", (emp_id,))
        row_data = c.fetchone()
        for idx, f in enumerate(EMPLOYEE_FIELDS):
            self.form_fields[f].setText(row_data[idx])
        self.selected_emp_id = emp_id
        self.load_attachments(emp_id)

    def load_attachments(self, emp_id):
        """
        Loads all attachments for the selected employee from the database into the attachments list widget.

        Clears the attachments list widget, then executes a SELECT query to retrieve all attachment records
        for the selected employee. For each record, adds an item to the attachments list widget and
        appends the record to the attachments attribute. If the attachment is a photo, sets the
        photo_path attribute to the attachment's filepath. Finally, calls the display_photo method to
        display the photo.

        :param emp_id: The ID of the selected employee.
        :type emp_id: int
        :return: None
        :rtype: NoneType
        """
        self.attach_list.clear()
        c = self.conn.cursor()
        c.execute("SELECT filename, filepath, is_photo FROM attachment WHERE employee_id=?", (emp_id,))
        self.attachments = []
        self.photo_path = None
        for fname, fpath, is_photo in c.fetchall():
            self.attach_list.addItem(fname)
            self.attachments.append((fname, fpath))
            if is_photo:
                self.photo_path = fpath
        self.display_photo()
        # Enable double-click to open attachment
        self.attach_list.itemDoubleClicked.connect(self.open_attachment)

    def open_attachment(self, item):
        """
        Opens the selected attachment file in the default application.
        """
        # Find the file path for the selected item
        fname = item.text()
        for name, path in self.attachments:
            if name == fname and os.path.exists(path):
                if sys.platform.startswith('darwin'):
                    os.system(f'open "{path}"')
                elif os.name == 'nt':
                    os.startfile(path)
                elif os.name == 'posix':
                    os.system(f'xdg-open "{path}"')
                break

    def display_photo(self):
        """
        Displays the selected employee's photo in the photo label.

        If the selected employee has a photo, scales it to 80x80 while maintaining aspect ratio,
        and sets it as the pixmap of the photo label. Otherwise, clears the photo label.

        :return: None
        :rtype: NoneType
        """
        if self.photo_path and os.path.exists(self.photo_path):
            pixmap = QPixmap(self.photo_path).scaled(80, 80, Qt.KeepAspectRatio)
            self.photo_label.setPixmap(pixmap)
        else:
            self.photo_label.clear()

    def upload_files(self):
        """
        Opens a file dialog for selecting files to upload as attachments for the selected employee.

        For each selected file, copies it to the employee's folder in the attachments directory and adds an item to the
        attachments list widget. Also appends the file record to the attachments attribute.
        If the selected employee ID is not None, inserts a new record into the attachment table
        with the selected employee ID, filename, filepath, filetype, upload_date, and is_photo set to 0.

        :return: None
        :rtype: NoneType
        """
        files, _ = QFileDialog.getOpenFileNames(self, "اختر ملفات")
        file_no = self.form_fields["file_no"].text()
        if not file_no:
            QMessageBox.critical(self, "خطأ", "يرجى إدخال رقم الملف أولاً")
            return
        emp_folder = get_employee_folder(file_no)
        for f in files:
            fname = os.path.basename(f)
            dest = os.path.join(emp_folder, fname)
            if not os.path.exists(dest):
                try:
                    with open(f, "rb") as src, open(dest, "wb") as dst:
                        dst.write(src.read())
                except Exception:
                    continue
            self.attach_list.addItem(fname)
            self.attachments.append((fname, dest))
            # Save to DB if editing existing employee
            if self.selected_emp_id:
                filetype, _ = mimetypes.guess_type(dest)
                upload_date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                c = self.conn.cursor()
                c.execute(
                    "INSERT INTO attachment (employee_id, filename, filepath, filetype, upload_date, is_photo) VALUES (?, ?, ?, ?, ?, ?)",
                    (self.selected_emp_id, fname, dest, filetype or '', upload_date, 0)
                )
                self.conn.commit()

    def upload_photo(self):
        """
        Opens a file dialog for selecting a photo to upload as an attachment for the selected employee.

        If a file is selected, copies it to the employee's folder in the attachments directory and sets the photo path attribute.
        Then calls the display_photo method to display the photo in the photo label.
        If the selected employee ID is not None, inserts a new record into the attachment table
        with the selected employee ID, filename, filepath, filetype, upload_date, and is_photo set to 1.

        :return: None
        :rtype: NoneType
        """
        f, _ = QFileDialog.getOpenFileName(self, "اختر صورة شخصية", "", "Images (*.png *.jpg *.jpeg)")
        if f:
            file_no = self.form_fields["file_no"].text()
            if not file_no:
                QMessageBox.critical(self, "خطأ", "يرجى إدخال رقم الملف أولاً")
                return
            emp_folder = get_employee_folder(file_no)
            fname = "photo_" + os.path.basename(f)
            dest = os.path.join(emp_folder, fname)
            with open(f, "rb") as src, open(dest, "wb") as dst:
                dst.write(src.read())
            self.photo_path = dest
            self.display_photo()
            # Save to DB if editing existing employee
            if self.selected_emp_id:
                filetype, _ = mimetypes.guess_type(dest)
                upload_date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                c = self.conn.cursor()
                c.execute(
                    "INSERT INTO attachment (employee_id, filename, filepath, filetype, upload_date, is_photo) VALUES (?, ?, ?, ?, ?, ?)",
                    (self.selected_emp_id, fname, dest, filetype or '', upload_date, 1)
                )
                self.conn.commit()

    def validate_employee_form(self, skip_id=None):
        """
        Validates the employee form fields.
        Returns (True, "") if valid, otherwise (False, error_message).
        skip_id: employee id to skip when checking uniqueness (for edit).
        """
        name = self.form_fields["name"].text().strip()
        file_no = self.form_fields["file_no"].text().strip()
        national_id = self.form_fields["national_id"].text().strip()
        phone = self.form_fields["phone"].text().strip()
        date_fields = ["grade_date", "hire_date", "birth_date"]

        # الاسم: مطلوب
        if not name:
            return False, "يرجى إدخال الاسم"

        # رقم الملف: مطلوب وفريد
        if not file_no:
            return False, "يرجى إدخال رقم الملف"
        c = self.conn.cursor()
        if skip_id:
            c.execute("SELECT id FROM employee WHERE file_no=? AND id!=?", (file_no, skip_id))
        else:
            c.execute("SELECT id FROM employee WHERE file_no=?", (file_no,))
        if c.fetchone():
            return False, "رقم الملف مسجل بالفعل"

        # الرقم القومي: مطلوب، 14 رقم، فريد
        if not national_id:
            return False, "يرجى إدخال الرقم القومي"
        if not (national_id.isdigit() and len(national_id) == 14):
            return False, "الرقم القومي يجب أن يكون 14 رقمًا"
        if skip_id:
            c.execute("SELECT id FROM employee WHERE national_id=? AND id!=?", (national_id, skip_id))
        else:
            c.execute("SELECT id FROM employee WHERE national_id=?", (national_id,))
        if c.fetchone():
            return False, "هذا الرقم القومي مسجل بالفعل."

        # رقم الهاتف: اختياري، لكن إذا أدخل يجب أن يكون أرقام فقط
        if phone and not phone.isdigit():
            return False, "رقم التليفون يجب أن يحتوي على أرقام فقط"

        # التواريخ: لا يمكن أن تكون في المستقبل
        today = datetime.date.today()
        for field in date_fields:
            date_str = self.form_fields[field].text().strip()
            if date_str:
                try:
                    # Try parsing as YYYY-MM-DD
                    date_val = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
                    if date_val > today:
                        return False, f"{AR_LABELS[field]} لا يمكن أن يكون في المستقبل"
                except Exception:
                    return False, f"صيغة التاريخ غير صحيحة في {AR_LABELS[field]} (يرجى استخدام YYYY-MM-DD)"

        return True, ""

    def add_employee(self):
        """
        Adds a new employee to the database, or updates the current one if already selected.
        Prevents duplicate records using 'national_id' and 'file_no' as unique identifiers.
        """
        vals = [normalize_arabic(self.form_fields[f].text()) for f in EMPLOYEE_FIELDS]
        # Validation
        valid, msg = self.validate_employee_form(self.selected_emp_id)
        if not valid:
            QMessageBox.critical(self, "خطأ", msg)
            return

        c = self.conn.cursor()

        # If an employee is selected, update instead of insert
        if self.selected_emp_id:
            c.execute(
                f"UPDATE employee SET {', '.join([f'{f}=?' for f in EMPLOYEE_FIELDS])} WHERE id=?",
                vals + [self.selected_emp_id]
            )
            # Remove old attachments and re-insert
            c.execute("DELETE FROM attachment WHERE employee_id=?", (self.selected_emp_id,))
            for fname, fpath in self.attachments:
                c.execute("INSERT INTO attachment (employee_id, filename, filepath, is_photo) VALUES (?, ?, ?, ?)",
                          (self.selected_emp_id, fname, fpath, 1 if fpath == self.photo_path else 0))
            self.conn.commit()
            self.load_employees()
            self.clear_form()
            QMessageBox.information(self, "تم", "تم تحديث بيانات الموظف بنجاح")
        else:
            # Insert new employee
            c.execute(
                f"INSERT INTO employee ({', '.join(EMPLOYEE_FIELDS)}) VALUES ({', '.join(['?']*len(EMPLOYEE_FIELDS))})",
                vals
            )
            emp_id = c.lastrowid
            for fname, fpath in self.attachments:
                c.execute("INSERT INTO attachment (employee_id, filename, filepath, is_photo) VALUES (?, ?, ?, ?)",
                          (emp_id, fname, fpath, 1 if fpath == self.photo_path else 0))
            self.conn.commit()
            self.load_employees()
            self.clear_form()
            QMessageBox.information(self, "تم", "تم إضافة الموظف بنجاح")

    def edit_employee(self):
        """
        Edits the selected employee in the database.

        If the selected employee ID is not None, updates the employee record in the employee table
        with the values from the form fields. Then deletes all attachments of the selected employee and
        inserts new attachments from the attachments list widget. Finally, commits the changes, reloads the
        employees, clears the form, and shows an information message box with a success message.
        """
        if not self.selected_emp_id:
            QMessageBox.critical(self, "خطأ", "يرجى اختيار موظف للتعديل")
            return
        # Validation
        valid, msg = self.validate_employee_form(self.selected_emp_id)
        if not valid:
            QMessageBox.critical(self, "خطأ", msg)
            return
        vals = [normalize_arabic(self.form_fields[f].text()) for f in EMPLOYEE_FIELDS]
        c = self.conn.cursor()
        c.execute(f"UPDATE employee SET {', '.join([f'{f}=?' for f in EMPLOYEE_FIELDS])} WHERE id=?", vals + [self.selected_emp_id])
        c.execute("DELETE FROM attachment WHERE employee_id=?", (self.selected_emp_id,))
        for fname, fpath in self.attachments:
            c.execute("INSERT INTO attachment (employee_id, filename, filepath, is_photo) VALUES (?, ?, ?, ?)",
                      (self.selected_emp_id, fname, fpath, 1 if fpath == self.photo_path else 0))
        self.conn.commit()
        self.load_employees()
        self.clear_form()
        QMessageBox.information(self, "تم", "تم تعديل بيانات الموظف")

    def delete_employee(self):
        """
        Deletes the selected employee from the database.

        If the selected employee ID is not None, prompts the user to confirm deletion.
        If the user confirms, deletes the employee record from the employee table and all attachments of the selected employee from the attachment table.
        Finally, commits the changes, reloads the employees, clears the form, and shows an information message box with a success message.
        """
        if not self.selected_emp_id:
            QMessageBox.critical(self, "خطأ", "يرجى اختيار موظف للحذف")
            return
        reply = QMessageBox.question(self, "تأكيد", "هل أنت متأكد من حذف الموظف؟", QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            c = self.conn.cursor()
            c.execute("DELETE FROM employee WHERE id=?", (self.selected_emp_id,))
            c.execute("DELETE FROM attachment WHERE employee_id=?", (self.selected_emp_id,))
            self.conn.commit()
            self.load_employees()
            self.clear_form()
            QMessageBox.information(self, "تم", "تم حذف الموظف")

    def clear_form(self):
        """
        Clears all form fields, attachments list widget, attachments list, photo path, photo label, and selected employee ID.

        :return: None
        :rtype: NoneType
        """
        for f in EMPLOYEE_FIELDS:
            self.form_fields[f].clear()
        self.attach_list.clear()
        self.attachments = []
        self.photo_path = None
        self.photo_label.clear()
        self.selected_emp_id = None

    def search_employees(self, search_text):
        """
        Searches for employees in the database by name, department, file_no, or national_id.

        Clears the table widget, then executes a SELECT query to retrieve all employee records
        where the name, department, file_no, or national_id matches the search text (case-insensitive).
        For each record, inserts a new row into the table widget and sets the values of the row
        according to the record's fields.

        :param search_text: The search text.
        :type search_text: str
        :return: None
        :rtype: NoneType
        """
        self.table.setRowCount(0)
        c = self.conn.cursor()
        # Normalize search text
        norm_search = normalize_arabic(search_text)
        query = f"""
            SELECT id, {', '.join(EMPLOYEE_FIELDS)} FROM employee
            WHERE name LIKE ?
               OR department LIKE ?
               OR file_no LIKE ?
               OR national_id LIKE ?
        """
        like_text = f"%{norm_search}%"
        c.execute(query, (like_text, like_text, like_text, like_text))
        for row in c.fetchall():
            row_idx = self.table.rowCount()
            self.table.insertRow(row_idx)
            for col_idx, val in enumerate(row[1:]):
                self.table.setItem(row_idx, col_idx, QTableWidgetItem(str(val)))
            self.table.setVerticalHeaderItem(row_idx, QTableWidgetItem(str(row[0])))

class ReportTab(QWidget):
    def __init__(self, conn):
        """
        Initializes the report tab.

        Sets the window title and layout. It also sets up the signals of the UI elements to their respective slots.

        :param conn: The database connection.
        :type conn: sqlite3.Connection
        """
        super().__init__()
        self.conn = conn
        layout = QVBoxLayout()
        lbl = QLabel("التقارير")
        lbl.setStyleSheet("font-size:20px; font-weight:bold; color:#1976d2;")
        layout.addWidget(lbl)
        btns_layout = QHBoxLayout()
        # self.btn_basic = QPushButton("تقرير الموظفين الأساسي")
        # self.btn_basic.clicked.connect(self.basic_report)
        # btns_layout.addWidget(self.btn_basic)
        self.btn_full = QPushButton("تقرير شامل للموظفين")
        self.btn_full.clicked.connect(self.full_report)
        btns_layout.addWidget(self.btn_full)
        self.btn_pdf = QPushButton("تصدير كـ PDF")
        self.btn_pdf.clicked.connect(self.export_pdf)
        btns_layout.addWidget(self.btn_pdf)
        layout.addLayout(btns_layout)
        self.report_text = QTextEdit()
        self.report_text.setReadOnly(True)
        layout.addWidget(self.report_text)
        self.setLayout(layout)

    def basic_report(self):
        """
        Generates a basic report of employees with their name, grade, job title, department, and phone number.

        Executes a SELECT query to retrieve all employee records with the specified fields.
        Then constructs a report string with the retrieved records and sets the report text area to the report string.

        :return: None
        :rtype: NoneType
        """
        c = self.conn.cursor()
        c.execute(f"SELECT name, grade, job_title, department, phone FROM employee")
        rows = c.fetchall()
        report = f"{' | '.join([AR_LABELS[f] for f in ['name','grade','job_title','department','phone']])}\n"
        report += "-"*80 + "\n"
        for row in rows:
            report += " | ".join(row) + "\n"
        self.report_text.setPlainText(report)

    def full_report(self):
        """
        Generates a full report of employees with all their fields.

        Executes a SELECT query to retrieve all employee records with all fields.
        Then constructs a report string with the retrieved records and sets the report text area to the report string.

        :return: None
        :rtype: NoneType
        """
        c = self.conn.cursor()
        c.execute(f"SELECT {', '.join(EMPLOYEE_FIELDS)} FROM employee")
        rows = c.fetchall()
        report = f"{' | '.join([AR_LABELS[f] for f in EMPLOYEE_FIELDS])}\n"
        report += "-"*120 + "\n"
        for row in rows:
            report += " | ".join([str(r) if r else "" for r in row]) + "\n"
        self.report_text.setPlainText(report)

    def export_pdf(self):
        """
        Exports the full report as a landscape A4 PDF, one employee per row, wrapped cells, Arabic support, RTL using WeasyPrint.

        Retrieves all employee records from the database, generates a default file name with the current date and time,
        asks the user for a file location using a file dialog, builds an HTML table with RTL and Arabic font, then saves the HTML to a PDF file using WeasyPrint.

        Shows an information message box with a success message if the export is successful, or a critical message box with an error message if an exception occurs during the export process.

        :return: None
        :rtype: NoneType
        """
        c = self.conn.cursor()
        c.execute(f"SELECT {', '.join(EMPLOYEE_FIELDS)} FROM employee")
        rows = c.fetchall()
        headers = [AR_LABELS[f] for f in EMPLOYEE_FIELDS]

        # Generate default file name with current date and time
        now = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        default_name = f"Employees {now}.pdf"

        # Ask user for file location
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "حفظ التقرير كـ PDF",
            default_name,
            "PDF Files (*.pdf)"
        )
        if not file_path:
            return

        # Build HTML table with RTL and Arabic font
        html = f"""
        <html lang="ar">
        <head>
            <meta charset="utf-8">
            <style>
                @font-face {{
                    font-family: 'Amiri';
                    src: url('Amiri-Regular.ttf');
                }}
                body {{
                    direction: rtl;
                    font-family: 'Amiri', 'Cairo', 'Tahoma', sans-serif;
                    font-size: 10px;
                }}
                table {{
                    border-collapse: collapse;
                    width: 100%;
                    table-layout: fixed;
                }}
                th, td {{
                    border: 1px solid #888;
                    padding: 6px 4px;
                    word-break: break-word;
                    vertical-align: top;
                    text-align: right;
                }}
                th {{
                    background: #b3d1f7;
                }}
            </style>
        </head>
        <body>
            <h2 style="text-align:center;">تقرير الموظفين</h2>
            <table dir="ltr">
                <thead>
                    <tr>
                        {''.join(f'<th>{h}</th>' for h in headers[::-1])}
                    </tr>
                </thead>
                <tbody>
        """
        for row in rows:
            html += "<tr>"
            for cell in row[::-1]:
                html += f"<td>{cell if cell else ''}</td>"
            html += "</tr>"
        html += """
                </tbody>
            </table>
        </body>
        </html>
        """

        # Save HTML to PDF using WeasyPrint
        try:
            css = CSS(string='''
                @page { size: A4 landscape; margin: 1cm; }
            ''')
            HTML(string=html, base_url=os.getcwd()).write_pdf(file_path, stylesheets=[css])
            QMessageBox.information(self, "تم", "تم تصدير التقرير بنجاح كملف PDF.")
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء تصدير التقرير: {e}")

if __name__ == "__main__":
    init_db()
    app = QApplication(sys.argv)
    window = MasarMainWindow()
    window.show()
    sys.exit(app.exec_())

"""
تعليمات البناء (Build Instructions):

1. تأكد من تثبيت المتطلبات:
   pip install pillow

2. لحزم التطبيق كملف exe:
   - ثبت pyinstaller: pip install pyinstaller
   - شغل الأمر:
     pyinstaller --onefile --windowed masar.py

   سيظهر الملف التنفيذي في مجلد dist.

3. تأكد من نسخ مجلد attachments بجانب الملف التنفيذي.

"""

# Copyright 2025 Shehab.Magdy.Eladl
# Licensed under the Apache License, Version 2.0 (see LICENSE file)