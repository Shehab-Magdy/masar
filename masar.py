import sys
import os
import sqlite3
import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget, QVBoxLayout, QLabel, QPushButton,
    QTableWidget, QTableWidgetItem, QLineEdit, QHBoxLayout, QFileDialog, QListWidget,
    QMessageBox, QTextEdit, QFormLayout, QSizePolicy, QGridLayout, QDateEdit, QListWidgetItem
)
from PyQt5.QtWidgets import QInputDialog
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt, pyqtSignal
from weasyprint import HTML, CSS
import mimetypes
import shutil
import base64
import calendar
import time
from pdf_bg_utils import process_bg_image


class ClickableLabel(QLabel):
    """A QLabel that emits a clicked signal when pressed."""
    clicked = pyqtSignal()

    def mousePressEvent(self, event):
        if event.button() == 1:
            self.clicked.emit()
        super().mousePressEvent(event)

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
    "personal_photo": "صورة شخصية",
    "retirement_date": "تاريخ المعاش"
    ,"insurance_doc": "وثيقة التامين"
    ,"serial": "مسلسل"
}

EMPLOYEE_FIELDS = [
    "name", "grade", "grade_date", "hire_date", "file_no", "qualification",
    "functional_group", "type_group", "job_title", "department", "current_work",
    "birth_date", "retirement_date", "insurance_no", "national_id", "address", "phone", "insurance_doc", "notes"
]

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
        self.setWindowTitle("مسار - منظومة العاملين المدنيين بالورش")
        self.setGeometry(100, 100, 1100, 700)
        # Set window icon
        from PyQt5.QtGui import QIcon
        self.setWindowIcon(QIcon(os.path.join(os.path.dirname(__file__), "masar.ico")))
        self.conn = sqlite3.connect(DB_FILE)
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)
        self.tabs.addTab(DashboardTab(self.conn), "الإحصائيات")
        self.tabs.addTab(EmployeeTab(self.conn), "الموظفين")
        # Correspondence tab (المراسلات)
        try:
            self.tabs.addTab(CorrespondenceTab(self.conn), "المراسلات")
        except NameError:
            # Class may be defined later in file; adding tab will work after import
            pass


class DashboardTab(QWidget):
    def __init__(self, conn):
        super().__init__()
        self.conn = conn
        layout = QVBoxLayout()
        self.lbl_title = QLabel("الإحصائيات")
        self.lbl_title.setStyleSheet("font-size:24px; font-weight:bold; color:#1976d2;")
        layout.addWidget(self.lbl_title)
        self.lbl_emp = QLabel()
        self.lbl_dept = QLabel()
        self.lbl_att = QLabel()
        self.lbl_retire_this_year = QLabel()
        layout.addWidget(self.lbl_emp)
        layout.addWidget(self.lbl_dept)
        layout.addWidget(self.lbl_att)
        layout.addWidget(self.lbl_retire_this_year)

        # --- Add table for employees retiring this year ---
        # self.retire_table = QTableWidget()
        # self.retire_table.setColumnCount(2)
        # self.retire_table.setHorizontalHeaderLabels(["الاسم", "رقم الملف"])
        # self.retire_table.setEditTriggers(QTableWidget.NoEditTriggers)
        # self.retire_table.setSelectionBehavior(QTableWidget.SelectRows)
        # self.retire_table.setSortingEnabled(False)
        # layout.addWidget(QLabel("الموظفون الذين تاريخ معاشهم في هذا العام:"))
        # layout.addWidget(self.retire_table)

        # --- Add refresh and print buttons side by side ---
        btns_layout = QHBoxLayout()
        self.btn_refresh = QPushButton("تحديث")
        self.btn_refresh.clicked.connect(self.refresh_counts)
        btns_layout.addWidget(self.btn_refresh)
        self.btn_print_retire = QPushButton("تصدير الموظفون الذين تاريخ معاشهم في هذا العام كـ PDF")
        self.btn_print_retire.clicked.connect(self.prompt_and_export_retire)
        btns_layout.addWidget(self.btn_print_retire)
        layout.addLayout(btns_layout)

        self.setLayout(layout)
        self.refresh_counts()

    def refresh_counts(self):
        c = self.conn.cursor()
        c.execute("SELECT COUNT(*) FROM employee")
        emp_count = c.fetchone()[0]
        c.execute("SELECT COUNT(DISTINCT department) FROM employee")
        dept_count = c.fetchone()[0]
        c.execute("SELECT COUNT(*) FROM attachment")
        att_count = c.fetchone()[0]
        current_year = datetime.date.today().year
        c.execute("""
            SELECT COUNT(*) FROM employee
            WHERE retirement_date IS NOT NULL
              AND retirement_date != ''
              AND substr(retirement_date, 1, 4) = ?
        """, (str(current_year),))
        retire_this_year = c.fetchone()[0]
        self.lbl_emp.setText(f"عدد الموظفين: {emp_count}")
        self.lbl_dept.setText(f"عدد الأقسام: {dept_count}")
        self.lbl_att.setText(f"عدد الملفات المرفوعة: {att_count}")
        self.lbl_retire_this_year.setText(f"عدد الموظفين الذين تاريخ معاشهم في هذا العام: {retire_this_year}")

        # --- Fill the retire_table with employees retiring this year ---
        # self.retire_table.setRowCount(0)
        # c.execute("""
        #     SELECT name, file_no FROM employee
        #     WHERE retirement_date IS NOT NULL
        #       AND retirement_date != ''
        #       AND substr(retirement_date, 1, 4) = ?
        # """, (str(current_year),))
        # for row_idx, (name, file_no) in enumerate(c.fetchall()):
        #     self.retire_table.insertRow(row_idx)
        #     self.retire_table.setItem(row_idx, 0, QTableWidgetItem(str(name)))
        #     self.retire_table.setItem(row_idx, 1, QTableWidgetItem(str(file_no)))

    def export_retire_pdf(self):
        """
        Export the full data of employees whose retirement date is in the current year as a PDF,
        using the same split-header, 9-columns-per-row, two-rows-per-employee design as export_filtered_pdf/export_pdf.
        """
        # Backwards-compatible wrapper: if called without months, export current-year retirees
        return self._export_retire_pdf_months(None)

    def prompt_and_export_retire(self):
        """
        Show an input dialog asking for the number of months (عدد الاشهر) and call the exporter.
        The entered number represents the next N months from the current month.
        """
        try:
            months, ok = QInputDialog.getInt(self, "إدخال عدد الأشهر", "من فضلك أدخل عدد الأشهر القادمة لحساب من يقترب تاريخ تقاعدهم:", value=1, min=1, max=120)
        except Exception:
            months, ok = QInputDialog.getInt(self, "إدخال عدد الأشهر", "من فضلك أدخل عدد الأشهر القادمة لحساب من يقترب تاريخ تقاعدهم:", value=1, min=1, max=120)
        if not ok:
            return
        self._export_retire_pdf_months(months)

    def _export_retire_pdf_months(self, months: int | None):
        """
        Export employees whose `retirement_date` is within the next `months` months.
        # Accept empty input (treat as fallback to current-year behavior), default to 6
        """
        default_val = 6
        text, ok = QInputDialog.getText(self, "إدخال عدد الأشهر", "من فضلك أدخل عدد الأشهر القادمة لحساب من يقترب تاريخ تقاعدهم:", text=str(default_val))
        if not ok:
            return
        text = text.strip()
        if text == "":
            # empty input -> use fallback behavior (months=None)
            self._export_retire_pdf_months(None)
            return
        # try parse integer
        try:
            months = int(text)
        except Exception:
            QMessageBox.warning(self, "تنبيه", "الرجاء إدخال رقم صحيح للأشهر أو اتركه فارغًا.")
            return
        if months <= 0:
            QMessageBox.warning(self, "تنبيه", "الرجاء إدخال عدد أكبر من صفر أو اتركه فارغًا.")
            return
        self._export_retire_pdf_months(months)
        headers2 = [AR_LABELS[f] for f in fields2]
        # add leading serial column label 'م'
        headers = ["م"] + [AR_LABELS[f] for f in EMPLOYEE_FIELDS]
        today = datetime.date.today()

        c = self.conn.cursor()
        # Fetch all employees that have a non-empty retirement_date, then filter in Python when months provided
        c.execute(f"SELECT {', '.join(EMPLOYEE_FIELDS)} FROM employee WHERE retirement_date IS NOT NULL AND retirement_date != ''")
        all_rows = c.fetchall()

        def parse_retirement_date(s: str):
            if not s:
                return None
            s = s.strip()
            for fmt in ("%Y-%m-%d", "%Y-%m", "%Y"):
                try:
                    dt = datetime.datetime.strptime(s, fmt)
                    # normalize to a date object (use first day if month/year only)
                    if fmt == "%Y":
                        return datetime.date(dt.year, 1, 1)
                    if fmt == "%Y-%m":
                        return datetime.date(dt.year, dt.month, 1)
                    return datetime.date(dt.year, dt.month, dt.day)
                except Exception:
                    continue
            return None

        def add_months(d: datetime.date, months_to_add: int) -> datetime.date:
            # Add months to a date while keeping day within month bounds
            total_month = d.month - 1 + months_to_add
            new_year = d.year + total_month // 12
            new_month = total_month % 12 + 1
            last_day = calendar.monthrange(new_year, new_month)[1]
            new_day = min(d.day, last_day)
            return datetime.date(new_year, new_month, new_day)

        rows = []
        if months is None:
            # fallback behavior: current year retirees (preserve previous behavior)
            current_year = today.year
            for row in all_rows:
                rd = parse_retirement_date(row[EMPLOYEE_FIELDS.index('retirement_date')])
                if rd and rd.year == current_year:
                    rows.append(row)
        else:
            end_date = add_months(today, months)
            for row in all_rows:
                rd = parse_retirement_date(row[EMPLOYEE_FIELDS.index('retirement_date')])
                if rd and today <= rd <= end_date:
                    rows.append(row)

        if not rows:
            QMessageBox.warning(self, "تنبيه", "لا يوجد بيانات لتصديرها.")
            return

        now = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        default_name = f"Employees_Retirement_{now}.pdf"
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "حفظ قائمة المعاش كـ PDF",
            default_name,
            "PDF Files (*.pdf)"
        )
        if not file_path:
            return
        
        # Prepare background image as base64 (only if file and config exist and are valid)
        bg_url = None
        first_line_header = ""
        second_line_header = ""
        bg_path = os.path.join(os.getcwd(), 'masar-bg.png')
        cfg_path = os.path.join(os.getcwd(), 'config.json')
        if os.path.isfile(cfg_path):
            try:
                import json
                with open(cfg_path, 'r', encoding='utf-8') as f:
                    cfg = json.load(f)
                first_line_header = cfg.get('firstLineHeader', "")
                second_line_header = cfg.get('secondLineHeader', "")
            except Exception:
                first_line_header = ""
                second_line_header = ""
        if os.path.isfile(bg_path) and os.path.isfile(cfg_path):
            try:
                bg_bytes = process_bg_image(bg_path, cfg_path)
                bg_b64 = base64.b64encode(bg_bytes).decode('utf-8')
                bg_url = f"data:image/png;base64,{bg_b64}"
            except Exception:
                bg_url = None
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
                    font-size: 9px;
                    {'background: url("'+bg_url+'"); background-size: contain; background-repeat: no-repeat; background-position: center center;' if bg_url else ''}
                }}
                table {{
                    border-collapse: collapse;
                    width: 100%;
                    margin-bottom: 20px;
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
                /* light rows should be transparent so the PDF background shows through */
                tr:nth-child(odd) {{
                    background-color: transparent;
                }}
                tr:nth-child(even) {{
                    background-color: #f2f2f2;
                }}
                @page {{
                    size: A4 landscape;
                    margin: 1cm 1cm 2cm 1cm; /* extra bottom margin for footer */
                    @bottom-center {{
                        content: counter(page) "/" counter(pages);
                        font-family: 'Amiri', 'Cairo', 'Tahoma', sans-serif;
                        font-size: 12px;
                        color: #444;
                    }}
                }}
            </style>
        </head>
        <body>
            <h2 style="text-align:center;">بيانات الخروج على المعاش</h2>
            <table dir="rtl">
                <thead>
                    <tr>
                        {''.join(f'<th>{h}</th>' for h in headers)}
                    </tr>
                </thead>
                <tbody>
        """

        for idx, emp in enumerate(rows):  # or employees
            row_class = "zebra1" if idx % 2 == 0 else "zebra2"
            serial = idx + 1
            html += f'<tr class="{row_class}"><td>{serial}</td>' + ''.join(
                f'<td>{emp[i] if i < len(emp) and emp[i] else ""}</td>' for i in range(len(EMPLOYEE_FIELDS))
            ) + '</tr>'

        html += """
                </tbody>
            </table>
        </body>
        </html>
        """

        try:
            css = CSS(string="""
                @page { 
                      size: A4 landscape; margin: 1cm 0.5cm 1.5cm 0.5cm;
                        @top-right {
                            content: '""" + first_line_header + """\\A""" + second_line_header + """';
                            font-size: 15px;
                            color: #1976d2;
                            text-align: right;
                            white-space: pre;
                      }
            """)
            HTML(string=html, base_url=os.getcwd()).write_pdf(file_path, stylesheets=[css])
            QMessageBox.information(self, "تم", "تم تصدير القائمة بنجاح كملف PDF.")
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء تصدير القائمة: {e}")

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
        search_layout.addWidget(self.search_field)
        search_layout.addWidget(QLabel("بحث:"))
        layout.addLayout(search_layout)
        main_layout = QHBoxLayout()
        self.table = QTableWidget()
        self.table.setColumnCount(len(EMPLOYEE_FIELDS))
        self.table.setHorizontalHeaderLabels([AR_LABELS[f] for f in EMPLOYEE_FIELDS])
        self.table.setSelectionBehavior(self.table.SelectRows)
        self.table.cellClicked.connect(self.on_row_select)
        self.table.setSortingEnabled(True)
        # Zebra striping for GUI table
        self.table.setAlternatingRowColors(True)
        # use transparent for the base/background rows so the UI background shows through
        self.table.setStyleSheet("QTableWidget {alternate-background-color: #f2f2f2; background-color: transparent;}")
        main_layout.addWidget(self.table)
        
        # Replace the QFormLayout with a QGridLayout for two columns
        grid_layout = QGridLayout()
        grid_layout.setAlignment(Qt.AlignmentFlag.AlignRight)
        grid_layout.setHorizontalSpacing(12)
        grid_layout.setVerticalSpacing(8)

        self.form_fields = {f: QLineEdit() for f in EMPLOYEE_FIELDS}
        self.attach_list = QListWidget()
        num_fields = len(EMPLOYEE_FIELDS)
        fields_per_col = num_fields // 2 + num_fields % 2  # 9 if 18 fields

        for idx, f in enumerate(EMPLOYEE_FIELDS):
            row = idx % fields_per_col
            col = (idx // fields_per_col) * 2
            label = QLabel(AR_LABELS[f])
            label.setMinimumWidth(0)
            label.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Preferred)
            self.form_fields[f].setAlignment(Qt.AlignmentFlag.AlignLeft)
            self.form_fields[f].setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
            grid_layout.addWidget(label, row, col)
            grid_layout.addWidget(self.form_fields[f], row, col + 1)

        # Attachments and buttons below the grid
        row_offset = fields_per_col
        grid_layout.addWidget(QLabel(AR_LABELS["attachments"]), row_offset, 0)
        grid_layout.addWidget(self.attach_list, row_offset, 1, 1, 3)
        attach_btns_layout = QHBoxLayout()
        self.btn_attach = QPushButton("رفع ملفات")
        self.btn_attach.clicked.connect(self.upload_files)
        attach_btns_layout.addWidget(self.btn_attach)
        self.btn_delete_attachment = QPushButton("حذف ملف")
        self.btn_delete_attachment.clicked.connect(self.delete_attachment)
        attach_btns_layout.addWidget(self.btn_delete_attachment)
        grid_layout.addLayout(attach_btns_layout, row_offset + 1, 1, 1, 3)

        # Photo upload and label
        grid_layout.addWidget(QLabel(AR_LABELS["personal_photo"]), row_offset + 2, 0)
        self.btn_photo = QPushButton("رفع صورة")
        self.btn_photo.clicked.connect(self.upload_photo)
        grid_layout.addWidget(self.btn_photo, row_offset + 2, 1)
        self.photo_label = QLabel()
        grid_layout.addWidget(self.photo_label, row_offset + 2, 2)

        # Action buttons
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
        # Add a button for printing/exporting the filtered list
        self.btn_export_filtered = QPushButton("تصدير النتائج كـ PDF")
        self.btn_export_filtered.clicked.connect(self.export_filtered_pdf)
        btns_layout.addWidget(self.btn_export_filtered)
        grid_layout.addLayout(btns_layout, row_offset + 3, 0, 1, 4)


        form_widget = QWidget()
        form_widget.setLayout(grid_layout)
        form_widget.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
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
        vh_item = self.table.verticalHeaderItem(row)
        if vh_item is None:
            return
        emp_id = vh_item.text()
        c = self.conn.cursor()
        c.execute(f"SELECT {', '.join(EMPLOYEE_FIELDS)} FROM employee WHERE id=?", (emp_id,))
        row_data = c.fetchone()
        for idx, f in enumerate(EMPLOYEE_FIELDS):
            self.form_fields[f].setText(row_data[idx])
        self.selected_emp_id = emp_id

        # --- Load attachments and set photo_path correctly ---
        self.attachments = []
        self.photo_path = None
        self.attach_list.clear()
        c.execute("SELECT filename, filepath, is_photo FROM attachment WHERE employee_id=?", (emp_id,))
        for fname, fpath, is_photo in c.fetchall():
            self.attach_list.addItem(fname)
            self.attachments.append((fname, fpath, is_photo))
            if is_photo:
                self.photo_path = fpath
        self.display_photo()
        # Enable double-click to open attachment
        self.attach_list.itemDoubleClicked.connect(self.open_attachment)

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
            self.attachments.append((fname, fpath, is_photo))
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
        for name, path, _ in self.attachments:
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
            pixmap = QPixmap(self.photo_path).scaled(80, 80, Qt.AspectRatioMode.KeepAspectRatio)
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
            orig_fname = os.path.basename(f)
            ext = os.path.splitext(orig_fname)[1]
            now_str = datetime.datetime.now().strftime("%Y-%m-%d")
            # Automatic renaming: originalName_yyyy-mm-dd.ext
            fname = f"{os.path.splitext(orig_fname)[0]}_{now_str}{ext}"
            dest = os.path.join(emp_folder, fname)
            # Ensure unique name if file exists
            counter = 1
            while os.path.exists(dest):
                fname = f"{os.path.splitext(orig_fname)[0]}_{now_str}_{counter}{ext}"
                dest = os.path.join(emp_folder, fname)
                counter += 1
            try:
                with open(f, "rb") as src, open(dest, "wb") as dst:
                    dst.write(src.read())
            except Exception:
                continue
            self.attach_list.addItem(fname)
            self.attachments.append((fname, dest, 0))  # 0 for normal file
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
            orig_fname = os.path.basename(f)
            ext = os.path.splitext(orig_fname)[1]
            now_str = datetime.datetime.now().strftime("%Y-%m-%d")
            # Automatic renaming: photo_yyyy-mm-dd.ext
            fname = f"photo_{now_str}{ext}"
            dest = os.path.join(emp_folder, fname)
            # Ensure unique name if file exists
            counter = 1
            while os.path.exists(dest):
                fname = f"photo_{now_str}_{counter}{ext}"
                dest = os.path.join(emp_folder, fname)
                counter += 1
            with open(f, "rb") as src, open(dest, "wb") as dst:
                dst.write(src.read())
            self.photo_path = dest
            self.display_photo()
            self.attach_list.addItem(fname)
            self.attachments.append((fname, dest, 1))  # 1 for photo
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
        insurance_no = self.form_fields["insurance_no"].text().strip()
        phone = self.form_fields["phone"].text().strip()
        date_fields = ["grade_date", "hire_date", "birth_date"]
        # Add retirement_date for format validation only
        retirement_date_str = self.form_fields["retirement_date"].text().strip()
        if retirement_date_str:
            try:
                datetime.datetime.strptime(retirement_date_str, "%Y-%m-%d")
            except Exception:
                return False, f"صيغة التاريخ غير صحيحة في {AR_LABELS['retirement_date']} (يرجى استخدام YYYY-MM-DD)"

        # الاسم: مطلوب
        if not name:
            return False, "يرجى إدخال الاسم"

        # رقم الملف: مطلوب وفريد ويجب أن يكون رقم صحيح
        if not file_no:
            return False, "يرجى إدخال رقم الملف"
        if not file_no.isdigit():
            return False, "رقم الملف يجب أن يكون أرقام فقط"
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

        # الرقم التأمينى: إذا أدخل يجب أن يكون أرقام فقط
        if insurance_no and not insurance_no.isdigit():
            return False, "الرقم التأميني يجب أن يحتوي على أرقام فقط"

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
            for fname, fpath, is_photo in self.attachments:
                c.execute(
                    "INSERT INTO attachment (employee_id, filename, filepath, is_photo) VALUES (?, ?, ?, ?)",
                    (self.selected_emp_id, fname, fpath, is_photo)
                )
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
            for fname, fpath, is_photo in self.attachments:
                c.execute(
                    "INSERT INTO attachment (employee_id, filename, filepath, is_photo) VALUES (?, ?, ?, ?)",
                    (emp_id, fname, fpath, is_photo)
                )
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
        for fname, fpath, is_photo in self.attachments:
            c.execute(
                "INSERT INTO attachment (employee_id, filename, filepath, is_photo) VALUES (?, ?, ?, ?)",
                (self.selected_emp_id, fname, fpath, is_photo)
            )
        self.conn.commit()
        self.load_employees()
        self.clear_form()
        QMessageBox.information(self, "تم", "تم تعديل بيانات الموظف")

    def delete_employee(self):
        """
        Deletes the selected employee from the database and removes their attachments folder from disk.

        If the selected employee ID is not None, prompts the user to confirm deletion.
        If the user confirms, deletes the employee record from the employee table and all attachments of the selected employee from the attachment table.
        Also deletes the employee's attachments folder from disk.
        Finally, commits the changes, reloads the employees, clears the form, and shows an information message box with a success message.
        """
        if not self.selected_emp_id:
            QMessageBox.critical(self, "خطأ", "يرجى اختيار موظف للحذف")
            return
        reply = QMessageBox.question(self, "تأكيد", "هل أنت متأكد من حذف الموظف؟", QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            # Get file_no before deleting the employee
            file_no = self.form_fields["file_no"].text().strip()
            c = self.conn.cursor()
            c.execute("DELETE FROM employee WHERE id=?", (self.selected_emp_id,))
            c.execute("DELETE FROM attachment WHERE employee_id=?", (self.selected_emp_id,))
            self.conn.commit()
            # Delete the employee's folder and all its contents
            if file_no:
                emp_folder = get_employee_folder(file_no)
                if os.path.exists(emp_folder):
                    try:
                        shutil.rmtree(emp_folder)
                    except Exception as e:
                        print(f"Error deleting folder {emp_folder}: {e}")
            self.load_employees()
            self.clear_form()
            QMessageBox.information(self, "تم", "تم حذف الموظف وجميع ملفاته بنجاح")

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

    def delete_attachment(self):
        """
        Deletes the selected attachment from the attachments list, database, and disk.
        """
        selected_items = self.attach_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "تنبيه", "يرجى اختيار مرفق للحذف")
            return
        item = selected_items[0]
        fname = item.text()
        # Find the attachment tuple
        for att in self.attachments:
            if att[0] == fname:
                fpath = att[1]
                break
        else:
            QMessageBox.warning(self, "تنبيه", "لم يتم العثور على الملف")
            return

        reply = QMessageBox.question(self, "تأكيد", f"هل أنت متأكد من حذف المرفق '{fname}'؟", QMessageBox.Yes | QMessageBox.No)
        if reply != QMessageBox.Yes:
            return

        # Remove from DB if employee is saved
        if self.selected_emp_id:
            c = self.conn.cursor()
            c.execute("DELETE FROM attachment WHERE employee_id=? AND filename=?", (self.selected_emp_id, fname))
            self.conn.commit()
        # Remove from disk
        if os.path.exists(fpath):
            try:
                os.remove(fpath)
            except Exception as e:
                print(f"Error deleting file {fpath}: {e}")
        # Remove from attachments list and UI
        self.attachments = [att for att in self.attachments if att[0] != fname]
        self.attach_list.takeItem(self.attach_list.row(item))
        # If it was the photo, clear photo
        if hasattr(self, "photo_path") and fpath == self.photo_path:
            self.photo_path = None
            self.display_photo()

    def export_filtered_pdf(self):
        """
        Exports the currently filtered list of employees in the table as a printable PDF report using WeasyPrint.
        The table has a split header (two rows of 9 columns), and each employee record is displayed in two rows:
        - First row: first 9 fields (with labels)
        - Second row: next 9 fields (with labels)
        """
        # Gather data from the table (only visible/filtered rows)

        rows = []
        for row_idx in range(self.table.rowCount()):
            row = []
            for col_idx in range(self.table.columnCount()):
                item = self.table.item(row_idx, col_idx)
                row.append(item.text() if item else "")
            rows.append(row)

        if not rows:
            QMessageBox.warning(self, "تنبيه", "لا يوجد بيانات لتصديرها.")
            return

        half = (len(EMPLOYEE_FIELDS) + 1) // 2
        fields1 = EMPLOYEE_FIELDS[:half]
        fields2 = EMPLOYEE_FIELDS[half:]
        # Add leading Arabic serial column 'م'
        headers = ["م"] + [AR_LABELS[f] for f in EMPLOYEE_FIELDS]


        now = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        default_name = f"Employees_{now}.pdf"
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "حفظ النتائج كـ PDF",
            default_name,
            "PDF Files (*.pdf)"
        )
        if not file_path:
            return
        
        # Prepare background image as base64 (only if file and config exist and are valid)
        bg_url = None
        first_line_header = ""
        second_line_header = ""
        bg_path = os.path.join(os.getcwd(), 'masar-bg.png')
        cfg_path = os.path.join(os.getcwd(), 'config.json')
        if os.path.isfile(cfg_path):
            try:
                import json
                with open(cfg_path, 'r', encoding='utf-8') as f:
                    cfg = json.load(f)
                first_line_header = cfg.get('firstLineHeader', "")
                second_line_header = cfg.get('secondLineHeader', "")
            except Exception:
                first_line_header = ""
                second_line_header = ""
        if os.path.isfile(bg_path) and os.path.isfile(cfg_path):
            try:
                bg_bytes = process_bg_image(bg_path, cfg_path)
                bg_b64 = base64.b64encode(bg_bytes).decode('utf-8')
                bg_url = f"data:image/png;base64,{bg_b64}"
            except Exception:
                bg_url = None
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
                    font-size: 9px;
                    {'background: url("'+bg_url+'"); background-size: contain; background-repeat: no-repeat; background-position: center center;' if bg_url else ''}
                }}
                table {{
                    border-collapse: collapse;
                    width: 100%;
                    margin-bottom: 20px;
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
                /* light rows transparent so background/print paper shows through */
                tr.zebra1 {{ background-color: transparent; }}
                tr.zebra2 {{ background-color: #f2f2f2; }}
                @page {{
                    size: A4 landscape;
                    margin: 1cm 1cm 2cm 1cm;
                    @bottom-center {{
                        content: counter(page) "/" counter(pages);
                        font-family: 'Amiri', 'Cairo', 'Tahoma', sans-serif;
                        font-size: 12px;
                        color: #444;
                    }}
                }}
            </style>
        </head>
        <body>
            <h2 style="text-align:center;">بيانات الموظفين المدنيين في الورش الرئيسية للطائرات</h2>
            <table dir="rtl">
                <thead>
                    <tr>
                        {''.join(f'<th>{h}</th>' for h in headers)}
                    </tr>
                </thead>
                <tbody>
        """

        for idx, emp in enumerate(rows):  # or employees
            row_class = "zebra1" if idx % 2 == 0 else "zebra2"
            serial = idx + 1
            html += f'<tr class="{row_class}"><td>{serial}</td>' + ''.join(
                f'<td>{emp[i] if i < len(emp) and emp[i] else ""}</td>' for i in range(len(EMPLOYEE_FIELDS))
            ) + '</tr>'

        html += """
                </tbody>
            </table>
        </body>
        </html>
        """

        try:
            css = CSS(string="""
                @page { 
                      size: A4 landscape; margin: 1cm 0.5cm 1.5cm 0.5cm;
                        @top-right {
                            content: '""" + first_line_header + """\\A""" + second_line_header + """';
                            font-size: 15px;
                            color: #1976d2;
                            text-align: right;
                            white-space: pre;
                      }
            """)
            HTML(string=html, base_url=os.getcwd()).write_pdf(file_path, stylesheets=[css])
            QMessageBox.information(self, "تم", "تم تصدير النتائج بنجاح كملف PDF.")
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء تصدير النتائج: {e}")

    def export_pdf(self):
        """
        Exports all employees as a printable PDF report using WeasyPrint.
        The table has a split header (two rows of 9 columns), and each employee record is displayed in two rows:
        - First row: first 9 fields (with labels)
        - Second row: next 9 fields (with labels)
        """
        c = self.conn.cursor()
        c.execute(f"SELECT {', '.join(EMPLOYEE_FIELDS)} FROM employee")
        employees = c.fetchall()
        if not employees:
            QMessageBox.warning(self, "تنبيه", "لا يوجد بيانات لتصديرها.")
            return

        half = 9
        fields1 = EMPLOYEE_FIELDS[:half]
        fields2 = EMPLOYEE_FIELDS[half:]
        # leading serial column
        headers = ["م"] + [AR_LABELS[f] for f in EMPLOYEE_FIELDS]
        headers2 = [AR_LABELS[f] for f in fields2]

        now = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "حفظ التقرير كـ PDF",
            f"employees_report_{now}.pdf",
            "PDF Files (*.pdf)"
        )
        if not file_path:
            return

        # Prepare background image as base64 (only if file and config exist and are valid)
        bg_url = None
        bg_path = os.path.join(os.getcwd(), 'masar-bg.png')
        cfg_path = os.path.join(os.getcwd(), 'config.json')
        first_line_header = ""
        second_line_header = ""
        if os.path.isfile(cfg_path):
            try:
                import json
                with open(cfg_path, 'r', encoding='utf-8') as f:
                    cfg = json.load(f)
                first_line_header = cfg.get('firstLineHeader', "")
                second_line_header = cfg.get('secondLineHeader', "")
            except Exception:
                first_line_header = ""
                second_line_header = ""
        if os.path.isfile(bg_path) and os.path.isfile(cfg_path):
            try:
                bg_bytes = process_bg_image(bg_path, cfg_path)
                bg_b64 = base64.b64encode(bg_bytes).decode('utf-8')
                bg_url = f"data:image/png;base64,{bg_b64}"
            except Exception:
                bg_url = None
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
                    font-size: 9px;
                    {'background: url("'+bg_url+'"); background-size: contain; background-repeat: no-repeat; background-position: center center;' if bg_url else ''}
                }}
                table {{
                    border-collapse: collapse;
                    width: 100%;
                    margin-bottom: 20px;
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
                /* light rows should be transparent so the PDF background shows through */
                tr:nth-child(odd) {{
                    background-color: transparent;
                }}
                tr:nth-child(even) {{
                    background-color: #f2f2f2;
                }}
                @page {{
                    size: A4 landscape;
                    margin: 1cm 1cm 1.5cm 1cm; /* extra bottom margin for footer */
                    @bottom-center {{
                        content: counter(page) "/" counter(pages);
                        font-family: 'Amiri', 'Cairo', 'Tahoma', sans-serif;
                        font-size: 12px;
                        color: #444;
                    }}
                }}
            </style>
        </head>
        <body>
            <h2 style="text-align:center;">بيانات الموظفين المدنيين في الورش الرئيسية للطائرات</h2>
            <table dir="rtl">
                <thead>
                    <tr>
                        {''.join(f'<th>{h}</th>' for h in headers)}
                    </tr>
                </thead>
                <tbody>
        """

        for idx, emp in enumerate(employees):  # or employees
            row_class = "zebra1" if idx % 2 == 0 else "zebra2"
            serial = idx + 1
            html += f'<tr class="{row_class}"><td>{serial}</td>' + ''.join(
                f'<td>{emp[i] if i < len(emp) and emp[i] else ""}</td>' for i in range(len(EMPLOYEE_FIELDS))
            ) + '</tr>'

        html += """
                </tbody>
            </table>
        </body>
        </html>
        """

        try:
            css = CSS(string="""
                @page { 
                      size: A4 landscape; margin: 1cm 0.5cm 1.5cm 0.5cm;
                        @top-right {
                            content: '{first_line_header.replace("'", "\\'")}\\A{second_line_header.replace("'", "\\'")}';
                            font-size: 15px;
                            color: #1976d2;
                            text-align: right;
                            white-space: pre;
                      }
            """)
            HTML(string=html, base_url=os.getcwd()).write_pdf(file_path, stylesheets=[css])
            QMessageBox.information(self, "تم", "تم تصدير التقرير بنجاح كملف PDF.")
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء تصدير التقرير: {e}")


class CorrespondenceTab(QWidget):
    """Tab for managing correspondence (المراسلات) with full CRUD, search, and PDF export."""
    def __init__(self, conn):
        super().__init__()
        self.conn = conn
        self.selected_id = None
        # temporary new images (file paths selected but not yet saved to DB)
        self.temp_images = []
        # existing attachments loaded from DB for the selected correspondence
        self.current_attachments = []  # list of tuples (id, filename, filepath)

        main_layout = QVBoxLayout()

        # Search panel
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("من التاريخ:"))
        self.search_from = QDateEdit()
        self.search_from.setCalendarPopup(True)
        self.search_from.setDisplayFormat("yyyy-MM-dd")
        self.search_from.setDate(datetime.date.today())
        search_layout.addWidget(self.search_from)

        search_layout.addWidget(QLabel("إلى التاريخ:"))
        self.search_to = QDateEdit()
        self.search_to.setCalendarPopup(True)
        self.search_to.setDisplayFormat("yyyy-MM-dd")
        self.search_to.setDate(datetime.date.today())
        search_layout.addWidget(self.search_to)

        search_layout.addWidget(QLabel("الموضوع:"))
        self.search_subject = QLineEdit()
        self.search_subject.setPlaceholderText("بحث في الموضوع...")
        search_layout.addWidget(self.search_subject)

        self.btn_search = QPushButton("بحث")
        self.btn_search.clicked.connect(self.search_entries)
        search_layout.addWidget(self.btn_search)

        self.btn_export_results = QPushButton("تصدير النتائج")
        self.btn_export_results.clicked.connect(self.export_results_pdf)
        search_layout.addWidget(self.btn_export_results)

        main_layout.addLayout(search_layout)

        # Main content layout: table and form
        content_layout = QHBoxLayout()

        # Table of correspondence
        self.table = QTableWidget()
        self.table.setColumnCount(8)
        headers = ["مسلسل", "رقم الفاكس", "التاريخ", "من", "إلى", "الموضوع", "الملاحظات", "الصورة"]
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setSelectionBehavior(self.table.SelectRows)
        self.table.cellClicked.connect(self.on_row_select)
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet("QTableWidget {alternate-background-color: #f2f2f2; background-color: transparent;}")
        content_layout.addWidget(self.table, 2)

        # Form layout
        form_layout = QGridLayout()
        row = 0
        form_layout.addWidget(QLabel("رقم الفاكس:"), row, 0)
        self.fax_number = QLineEdit()
        form_layout.addWidget(self.fax_number, row, 1)
        row += 1

        form_layout.addWidget(QLabel("التاريخ:"), row, 0)
        self.fax_date = QDateEdit()
        self.fax_date.setCalendarPopup(True)
        self.fax_date.setDisplayFormat("yyyy-MM-dd")
        self.fax_date.setDate(datetime.date.today())
        form_layout.addWidget(self.fax_date, row, 1)
        row += 1

        form_layout.addWidget(QLabel("من:"), row, 0)
        self.from_person = QLineEdit()
        form_layout.addWidget(self.from_person, row, 1)
        row += 1

        form_layout.addWidget(QLabel("إلى:"), row, 0)
        self.to_person = QLineEdit()
        form_layout.addWidget(self.to_person, row, 1)
        row += 1

        form_layout.addWidget(QLabel("الموضوع:"), row, 0)
        self.subject = QLineEdit()
        form_layout.addWidget(self.subject, row, 1)
        row += 1

        form_layout.addWidget(QLabel("الملاحظات:"), row, 0)
        self.notes = QTextEdit()
        form_layout.addWidget(self.notes, row, 1)
        row += 1

        form_layout.addWidget(QLabel("رفع صور الفاكس:"), row, 0)
        btn_img = QPushButton("اختر صور")
        btn_img.clicked.connect(self.browse_image)
        form_layout.addWidget(btn_img, row, 1)
        # list widget showing selected and existing images
        from PyQt5.QtWidgets import QListWidgetItem
        self.fax_images_list = QListWidget()
        self.fax_images_list.setFixedHeight(90)
        form_layout.addWidget(self.fax_images_list, row, 2)
        # show thumbnail when clicking list items
        self.fax_images_list.itemClicked.connect(self.on_image_item_clicked)
        row += 1
        # remove selected image button
        btn_remove_img = QPushButton("حذف الصورة المحددة")
        btn_remove_img.clicked.connect(self.remove_selected_image)
        form_layout.addWidget(btn_remove_img, row, 2)
        row += 1
        # Thumbnail display
        form_layout.addWidget(QLabel("معاينة الصورة:"), row, 0)
        self.thumbnail = ClickableLabel()
        self.thumbnail.setFixedSize(220, 140)
        self.thumbnail.setStyleSheet("border: 1px solid #ccc; background: #fff;")
        self.thumbnail.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.thumbnail.clicked.connect(self.open_current_thumbnail)
        form_layout.addWidget(self.thumbnail, row, 1, 1, 2)
        row += 1

        # Action buttons
        btns = QHBoxLayout()
        self.btn_add = QPushButton("إضافة")
        self.btn_add.clicked.connect(self.add_entry)
        btns.addWidget(self.btn_add)
        self.btn_edit = QPushButton("تعديل")
        self.btn_edit.clicked.connect(self.edit_entry)
        btns.addWidget(self.btn_edit)
        self.btn_delete = QPushButton("حذف")
        self.btn_delete.clicked.connect(self.delete_entry)
        btns.addWidget(self.btn_delete)
        self.btn_clear = QPushButton("مسح")
        self.btn_clear.clicked.connect(self.clear_form)
        btns.addWidget(self.btn_clear)

        form_layout.addLayout(btns, row, 0, 1, 3)

        form_widget = QWidget()
        form_widget.setLayout(form_layout)
        content_layout.addWidget(form_widget, 1)

        main_layout.addLayout(content_layout)

        self.setLayout(main_layout)
        self.load_entries()

    def browse_image(self):
        files, _ = QFileDialog.getOpenFileNames(self, "اختر صور الفاكس", "", "Images (*.png *.jpg *.jpeg)")
        if not files:
            return
        for f in files:
            if f:
                self.temp_images.append(f)
                item = QListWidgetItem(os.path.basename(f))
                item.setData(Qt.UserRole, ("temp", f))
                self.fax_images_list.addItem(item)
        # show thumbnail of the last selected image
        if files:
            self._show_thumbnail_from_path(files[-1])

    def clear_form(self):
        self.selected_id = None
        self.fax_number.clear()
        self.fax_date.setDate(datetime.date.today())
        self.from_person.clear()
        self.to_person.clear()
        self.subject.clear()
        self.notes.clear()
        self.temp_images = []
        self.current_attachments = []
        self.fax_images_list.clear()

    def load_entries(self, where_clause: str = "", params: tuple = ()): 
        c = self.conn.cursor()
        q = "SELECT id, fax_number, fax_date, from_person, to_person, subject, notes, created_at FROM correspondence"
        if where_clause:
            q += " WHERE " + where_clause
        q += " ORDER BY fax_date DESC"
        c.execute(q, params)
        rows = c.fetchall()
        self.table.setRowCount(0)
        for idx, row in enumerate(rows):
            rid = row[0]
            self.table.insertRow(idx)
            # serial
            self.table.setItem(idx, 0, QTableWidgetItem(str(idx + 1)))
            self.table.setItem(idx, 1, QTableWidgetItem(str(row[1] or "")))
            self.table.setItem(idx, 2, QTableWidgetItem(str(row[2] or "")))
            self.table.setItem(idx, 3, QTableWidgetItem(str(row[3] or "")))
            self.table.setItem(idx, 4, QTableWidgetItem(str(row[4] or "")))
            self.table.setItem(idx, 5, QTableWidgetItem(str(row[5] or "")))
            self.table.setItem(idx, 6, QTableWidgetItem(str(row[6] or "")))
            # show count of images for the correspondence
            c2 = self.conn.cursor()
            c2.execute("SELECT COUNT(*) FROM correspondence_attachment WHERE correspondence_id=?", (rid,))
            cnt = c2.fetchone()[0]
            self.table.setItem(idx, 7, QTableWidgetItem(str(cnt)))
            self.table.setVerticalHeaderItem(idx, QTableWidgetItem(str(rid)))

    def on_row_select(self, row, col):
        vh = self.table.verticalHeaderItem(row)
        if vh is None:
            return
        rid = vh.text()
        c = self.conn.cursor()
        c.execute("SELECT id, fax_number, fax_date, from_person, to_person, subject, notes FROM correspondence WHERE id=?", (rid,))
        r = c.fetchone()
        if not r:
            return
        self.selected_id = r[0]
        self.fax_number.setText(r[1] or "")
        try:
            if r[2]:
                dt = datetime.datetime.strptime(r[2], "%Y-%m-%d").date()
                self.fax_date.setDate(dt)
        except Exception:
            pass
        self.from_person.setText(r[3] or "")
        self.to_person.setText(r[4] or "")
        self.subject.setText(r[5] or "")
        self.notes.setPlainText(r[6] or "")
        # load attachments from DB
        self.current_attachments = []
        self.temp_images = []
        self.fax_images_list.clear()
        c2 = self.conn.cursor()
        c2.execute("SELECT id, filename, filepath FROM correspondence_attachment WHERE correspondence_id=?", (self.selected_id,))
        for att in c2.fetchall():
            att_id, fname, fpath = att
            self.current_attachments.append((att_id, fname, fpath))
            item = QListWidgetItem(fname)
            item.setData(Qt.UserRole, ("stored", att_id, fpath))
            self.fax_images_list.addItem(item)
        # show thumbnail of first attachment if available
        if self.current_attachments:
            first_path = self.current_attachments[0][2]
            if first_path:
                self._show_thumbnail_from_path(first_path)

    def _validate_form(self):
        fax_no = self.fax_number.text().strip()
        subj = self.subject.text().strip()
        if not fax_no:
            return False, "رقم الفاكس مطلوب"
        if not subj:
            return False, "الموضوع مطلوب"
        # date validation
        d = self.fax_date.date().toPyDate()
        if d > datetime.date.today():
            return False, "التاريخ لا يمكن أن يكون في المستقبل"
        return True, ""

    def _save_image(self, fax_number: str) -> str | None:
        # legacy single-image saver: keep for compatibility (not used)
        if not self.temp_images:
            return None
        faxes_dir = os.path.join(ATTACHMENTS_DIR, 'Faxes')
        if not os.path.exists(faxes_dir):
            os.makedirs(faxes_dir)
        # save all temp images and return the first saved path (or None)
        saved_paths = []
        for idx, src_path in enumerate(self.temp_images):
            try:
                ext = os.path.splitext(src_path)[1]
                fname = f"{fax_number}_{int(time.time())}_{idx}{ext}"
                dest = os.path.join(faxes_dir, fname)
                with open(src_path, 'rb') as src, open(dest, 'wb') as dst:
                    dst.write(src.read())
                saved_paths.append(dest)
            except Exception:
                continue
        return saved_paths[0] if saved_paths else None

    def _save_images_and_create_attachments(self, correspondence_id: int, fax_number: str):
        """Save all temp images to disk and create rows in correspondence_attachment."""
        if not self.temp_images:
            return
        faxes_dir = os.path.join(ATTACHMENTS_DIR, 'Faxes')
        if not os.path.exists(faxes_dir):
            os.makedirs(faxes_dir)
        c = self.conn.cursor()
        for idx, src_path in enumerate(self.temp_images):
            try:
                ext = os.path.splitext(src_path)[1]
                fname = f"{fax_number}_{correspondence_id}_{int(time.time())}_{idx}{ext}"
                dest = os.path.join(faxes_dir, fname)
                with open(src_path, 'rb') as src, open(dest, 'wb') as dst:
                    dst.write(src.read())
                upload_date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                c.execute("INSERT INTO correspondence_attachment (correspondence_id, filename, filepath, upload_date) VALUES (?, ?, ?, ?)",
                          (correspondence_id, fname, dest, upload_date))
            except Exception:
                continue
        self.conn.commit()
        # clear temp images after saving
        self.temp_images = []
        # refresh attachments list if this correspondence is currently selected
        if self.selected_id == correspondence_id:
            # reload attachments into list widget
            self.fax_images_list.clear()
            c2 = self.conn.cursor()
            c2.execute("SELECT id, filename, filepath FROM correspondence_attachment WHERE correspondence_id=?", (self.selected_id,))
            for att in c2.fetchall():
                att_id, fname, fpath = att
                item = QListWidgetItem(fname)
                item.setData(Qt.UserRole, ("stored", att_id, fpath))
                self.fax_images_list.addItem(item)
        # ensure thumbnail cleared/updated
        if self.fax_images_list.count() == 0:
            self.thumbnail.clear()
        else:
            # show last saved image
            last = self.fax_images_list.item(self.fax_images_list.count() - 1)
            if last:
                self.on_image_item_clicked(last)

    def add_entry(self):
        ok, msg = self._validate_form()
        if not ok:
            QMessageBox.warning(self, "خطأ في البيانات", msg)
            return
        fax_no = self.fax_number.text().strip()
        fax_date = self.fax_date.date().toString("yyyy-MM-dd")
        from_p = normalize_arabic(self.from_person.text().strip())
        to_p = normalize_arabic(self.to_person.text().strip())
        subj = normalize_arabic(self.subject.text().strip())
        notes = self.notes.toPlainText().strip()
        created_at = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        c = self.conn.cursor()
        # insert correspondence record first (image attachments will be stored in correspondence_attachment)
        c.execute("INSERT INTO correspondence (fax_number, fax_date, from_person, to_person, subject, notes, image_path, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                  (fax_no, fax_date, from_p, to_p, subj, notes, '', created_at))
        corr_id = c.lastrowid
        # save any temp images and create attachment rows
        self._save_images_and_create_attachments(corr_id, fax_no)
        self.conn.commit()
        self.clear_form()
        self.load_entries()

    def edit_entry(self):
        if not self.selected_id:
            QMessageBox.warning(self, "تحذير", "الرجاء اختيار سجل للتعديل")
            return
        ok, msg = self._validate_form()
        if not ok:
            QMessageBox.warning(self, "خطأ في البيانات", msg)
            return
        fax_no = self.fax_number.text().strip()
        fax_date = self.fax_date.date().toString("yyyy-MM-dd")
        from_p = normalize_arabic(self.from_person.text().strip())
        to_p = normalize_arabic(self.to_person.text().strip())
        subj = normalize_arabic(self.subject.text().strip())
        notes = self.notes.toPlainText().strip()
        c = self.conn.cursor()
        c.execute("UPDATE correspondence SET fax_number=?, fax_date=?, from_person=?, to_person=?, subject=?, notes=? WHERE id=?",
                  (fax_no, fax_date, from_p, to_p, subj, notes, self.selected_id))
        # save any newly selected temp images and create attachment rows
        self._save_images_and_create_attachments(self.selected_id, fax_no)
        self.conn.commit()
        self.clear_form()
        self.load_entries()

    def delete_entry(self):
        if not self.selected_id:
            QMessageBox.warning(self, "تحذير", "الرجاء اختيار سجل للحذف")
            return
        ok = QMessageBox.question(self, "تأكيد", "هل تريد حذف السجل المحدد؟")
        if ok != QMessageBox.Yes:
            return
        c = self.conn.cursor()
        # remove attachment files and rows
        c.execute("SELECT filepath FROM correspondence_attachment WHERE correspondence_id=?", (self.selected_id,))
        for (fp,) in c.fetchall():
            try:
                if fp and os.path.exists(fp):
                    os.remove(fp)
            except Exception:
                pass
        c.execute("DELETE FROM correspondence_attachment WHERE correspondence_id=?", (self.selected_id,))
        # remove correspondence row
        c.execute("DELETE FROM correspondence WHERE id=?", (self.selected_id,))
        self.conn.commit()
        self.clear_form()
        self.load_entries()

    def search_entries(self):
        clauses = []
        params = []
        # date range
        from_d = self.search_from.date().toPyDate()
        to_d = self.search_to.date().toPyDate()
        if from_d and to_d:
            # ensure from_d <= to_d
            if from_d > to_d:
                QMessageBox.warning(self, "خطأ", "نطاق التواريخ غير صالح")
                return
            clauses.append("fax_date BETWEEN ? AND ?")
            params.extend([from_d.strftime("%Y-%m-%d"), to_d.strftime("%Y-%m-%d")])
        subj = self.search_subject.text().strip()
        if subj:
            clauses.append("subject LIKE ?")
            params.append(f"%{subj}%")
        where = " AND ".join(clauses)
        self.load_entries(where, tuple(params))

    def remove_selected_image(self):
        """Remove selected image from the list. If it's a stored attachment, delete from DB and disk."""
        item = self.fax_images_list.currentItem()
        if not item:
            return
        data = item.data(Qt.UserRole)
        if not data:
            # just remove from widget
            self.fax_images_list.takeItem(self.fax_images_list.currentRow())
            return
        if data[0] == "temp":
            # remove from temp_images list
            path = data[1]
            try:
                self.temp_images.remove(path)
            except ValueError:
                pass
            self.fax_images_list.takeItem(self.fax_images_list.currentRow())
            return
        if data[0] == "stored":
            att_id = data[1]
            fpath = data[2]
            # confirm deletion
            ok = QMessageBox.question(self, "تأكيد", "هل تريد حذف هذه الصورة من السجل؟")
            if ok != QMessageBox.Yes:
                return
            c = self.conn.cursor()
            try:
                if fpath and os.path.exists(fpath):
                    os.remove(fpath)
            except Exception:
                pass
            c.execute("DELETE FROM correspondence_attachment WHERE id=?", (att_id,))
            self.conn.commit()
            # remove from internal list and widget
            for i in range(self.fax_images_list.count()):
                it = self.fax_images_list.item(i)
                if it is None:
                    continue
                d = it.data(Qt.UserRole)
                if d and d[0] == "stored" and d[1] == att_id:
                    self.fax_images_list.takeItem(i)
                    break
        # clear thumbnail if no items left
        if self.fax_images_list.count() == 0:
            self.thumbnail.clear()

    def on_image_item_clicked(self, item):
        if item is None:
            return
        data = item.data(Qt.UserRole)
        if not data:
            return
        if data[0] == "temp":
            path = data[1]
        elif data[0] == "stored":
            path = data[2]
        else:
            path = None
        if path:
            self._show_thumbnail_from_path(path)

    def _show_thumbnail_from_path(self, path: str):
        try:
            if not path or not os.path.exists(path):
                self.thumbnail.setText("لا توجد صورة")
                self.current_thumbnail_path = None
                return
            pix = QPixmap(path)
            if pix.isNull():
                self.thumbnail.setText("لا يمكن عرض الصورة")
                self.current_thumbnail_path = None
                return
            scaled = pix.scaled(self.thumbnail.width(), self.thumbnail.height(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.thumbnail.setPixmap(scaled)
            self.current_thumbnail_path = path
        except Exception:
            self.thumbnail.setText("خطأ في عرض الصورة")
            self.current_thumbnail_path = None

    def open_current_thumbnail(self):
        path = getattr(self, 'current_thumbnail_path', None)
        if not path:
            return
        try:
            if sys.platform.startswith('darwin'):
                os.system(f'open "{path}"')
            elif os.name == 'nt':
                os.startfile(path)
            elif os.name == 'posix':
                os.system(f'xdg-open "{path}"')
        except Exception:
            QMessageBox.warning(self, "خطأ", "تعذر فتح الملف في التطبيق الافتراضي")

    def export_results_pdf(self):
        # reuse search to get rows
        clauses = []
        params = []
        from_d = self.search_from.date().toPyDate()
        to_d = self.search_to.date().toPyDate()
        if from_d and to_d:
            if from_d > to_d:
                QMessageBox.warning(self, "خطأ", "نطاق التواريخ غير صالح")
                return
            clauses.append("fax_date BETWEEN ? AND ?")
            params.extend([from_d.strftime("%Y-%m-%d"), to_d.strftime("%Y-%m-%d")])
        subj = self.search_subject.text().strip()
        if subj:
            clauses.append("subject LIKE ?")
            params.append(f"%{subj}%")
        where = " AND ".join(clauses)
        c = self.conn.cursor()
        q = "SELECT fax_number, fax_date, from_person, to_person, subject, notes, image_path FROM correspondence"
        if where:
            q += " WHERE " + where
        q += " ORDER BY fax_date DESC"
        c.execute(q, tuple(params))
        rows = c.fetchall()
        if not rows:
            QMessageBox.warning(self, "تنبيه", "لا توجد نتائج للتصدير")
            return
        now = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        file_path, _ = QFileDialog.getSaveFileName(self, "حفظ المراسلات كـ PDF", f"correspondence_{now}.pdf", "PDF Files (*.pdf)")
        if not file_path:
            return

        # prepare background
        bg_url = None
        bg_path = os.path.join(os.getcwd(), 'masar-bg.png')
        cfg_path = os.path.join(os.getcwd(), 'config.json')
        if os.path.isfile(bg_path) and os.path.isfile(cfg_path):
            try:
                bg_bytes = process_bg_image(bg_path, cfg_path)
                bg_b64 = base64.b64encode(bg_bytes).decode('utf-8')
                bg_url = f"data:image/png;base64,{bg_b64}"
            except Exception:
                bg_url = None

        # build html
        html = f"""
        <html lang="ar">
        <head>
          <meta charset="utf-8">
          <style>
            @font-face {{ font-family: 'Amiri'; src: url('Amiri-Regular.ttf'); }}
            body {{ direction: rtl; font-family: 'Amiri', 'Cairo', 'Tahoma', sans-serif; font-size: 11px; {'background: url("'+bg_url+'") no-repeat center center; background-size: contain;' if bg_url else ''} }}
            table {{ border-collapse: collapse; width: 100%; table-layout: fixed; word-wrap: break-word; }}
            th, td {{ border: 1px solid #888; padding: 6px; vertical-align: top; text-align: right; word-break: break-word; white-space: pre-wrap; }}
            th {{ background: #b3d1f7; }}
            tr:nth-child(odd) {{ background-color: transparent; }}
            tr:nth-child(even) {{ background-color: #f2f2f2; }}
            @page {{ size: A4 portrait; margin: 1cm; @bottom-center {{ content: counter(page) "/" counter(pages); }} }}
          </style>
        </head>
        <body>
          <h2 style="text-align:center;">قائمة المراسلات</h2>
          <table dir="rtl">
            <thead>
              <tr>
                <th>مسلسل</th>
                <th>رقم الفاكس</th>
                <th>التاريخ</th>
                <th>من</th>
                <th>إلى</th>
                <th>الموضوع</th>
                <th>الملاحظات</th>
              </tr>
            </thead>
            <tbody>
        """

        for idx, r in enumerate(rows):
            serial = idx + 1
            fax_no, fax_date, from_p, to_p, subj, notes, img = r
            html += f"<tr><td>{serial}</td><td>{fax_no or ''}</td><td>{fax_date or ''}</td><td>{from_p or ''}</td><td>{to_p or ''}</td><td>{(subj or '')}</td><td>{(notes or '').replace('\n','<br/>')}</td></tr>"

        html += """
            </tbody>
          </table>
        </body>
        </html>
        """

        try:
            css = CSS(string='@page { size: A4 portrait; margin: 1cm; }')
            HTML(string=html, base_url=os.getcwd()).write_pdf(file_path, stylesheets=[css])
            QMessageBox.information(self, "تم", "تم تصدير النتائج بنجاح كملف PDF.")
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء التصدير: {e}")


if __name__ == "__main__":
    init_db()
    app = QApplication(sys.argv)
    window = MasarMainWindow()
    window.showMaximized()
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