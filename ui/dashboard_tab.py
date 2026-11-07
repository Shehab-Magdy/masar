# ui/dashboard_tab.py
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QTableWidget, QPushButton, QHBoxLayout, QTableWidgetItem, QInputDialog, QMessageBox, QDialog, QFileDialog
import datetime
import os
import zipfile
from utils import constants
from reports.retire_export import export_retire_pdf
from ui.dialogs.PrintOptionsDialog import PrintOptionsDialog

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

        # Table for employee count by department
        self.dept_table = QTableWidget()
        self.dept_table.setColumnCount(2)
        self.dept_table.setHorizontalHeaderLabels(["القسم", "عدد الموظفين"])
        self.dept_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.dept_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.dept_table.setSortingEnabled(False)
        layout.addWidget(self.dept_table)

        # Table for employees retiring this year
        self.retire_table = QTableWidget()
        self.retire_table.setColumnCount(3)
        self.retire_table.setHorizontalHeaderLabels(["م", "الاسم", "رقم الملف"])
        self.retire_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.retire_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.retire_table.setSortingEnabled(False)
        layout.addWidget(self.retire_table)

        # Refresh and export buttons
        btns_layout = QHBoxLayout()
        self.btn_refresh = QPushButton("تحديث")
        self.btn_refresh.clicked.connect(self.refresh_counts)
        btns_layout.addWidget(self.btn_refresh)
        self.btn_print_retire = QPushButton("تصدير الموظفين الذين تاريخ معاشهم اقترب كـ PDF")
        self.btn_print_retire.clicked.connect(self.on_print_retire_clicked)
        btns_layout.addWidget(self.btn_print_retire)
        self.btn_backup = QPushButton("نسخة احتياطية")
        self.btn_backup.clicked.connect(self.backup_data)
        btns_layout.addWidget(self.btn_backup)
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
            WHERE retirement_date IS NOT NULL AND retirement_date != '' AND substr(retirement_date, 1, 4) = ?
        """, (str(current_year),))
        retire_this_year = c.fetchone()[0]
        self.lbl_emp.setText(f"عدد الموظفين: {emp_count}")
        self.lbl_dept.setText(f"عدد الأقسام: {dept_count}")
        self.lbl_att.setText(f"عدد الملفات المرفوعة: {att_count}")
        self.lbl_retire_this_year.setText(f"عدد الموظفين الذين تاريخ معاشهم في هذا العام: {retire_this_year}")

        # Fill the department breakdown table
        self.dept_table.setRowCount(0)
        c.execute("""
            SELECT department, COUNT(*) FROM employee
            GROUP BY department ORDER BY COUNT(*) DESC
        """)
        for row_idx, (department, count) in enumerate(c.fetchall()):
            self.dept_table.insertRow(row_idx)
            self.dept_table.setItem(row_idx, 0, QTableWidgetItem(str(department or "")))
            self.dept_table.setItem(row_idx, 1, QTableWidgetItem(str(count)))

        # Fill the retire_table with employees retiring this year
        self.retire_table.setRowCount(0)
        c.execute("""
            SELECT name, file_no FROM employee
            WHERE retirement_date IS NOT NULL AND retirement_date != '' AND substr(retirement_date, 1, 4) = ?
        """, (str(current_year),))
        for row_idx, (name, file_no) in enumerate(c.fetchall()):
            self.retire_table.insertRow(row_idx)
            self.retire_table.setItem(row_idx, 0, QTableWidgetItem(str(row_idx + 1)))
            self.retire_table.setItem(row_idx, 1, QTableWidgetItem(str(name)))
            self.retire_table.setItem(row_idx, 2, QTableWidgetItem(str(file_no)))

    def on_print_retire_clicked(self):
        """
        Show a dialog to get the number of upcoming months, validate, and call export_retire_pdf(months).
        """
        # Show print options dialog
        print_dialog = PrintOptionsDialog(self)
        if print_dialog.exec_() != QDialog.Accepted:
            return
        orientation = "landscape" if print_dialog.is_landscape() else "portrait"

        months, ok = QInputDialog.getInt(
            self,
            "إدخال عدد الأشهر",
            "من فضلك أدخل عدد الأشهر القادمة:",
            value=6, min=1, max=120
        )
        if not ok:
            return  # User cancelled
        if months < 1:
            QMessageBox.warning(self, "تنبيه", "يرجى إدخال رقم موجب أكبر من الصفر لعدد الأشهر.")
            return
        export_retire_pdf(self.conn, months, parent=self, orientation=orientation)

    def backup_data(self):
        """
        Zips the database file and the attachments directory into a single
        timestamped archive at a location chosen by the user.
        """
        now = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        default_name = f"Masar_Backup_{now}.zip"
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "حفظ نسخة احتياطية",
            default_name,
            "Zip Files (*.zip)"
        )
        if not file_path:
            return
        try:
            with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                if os.path.isfile(constants.DB_FILE):
                    zf.write(constants.DB_FILE, arcname=os.path.basename(constants.DB_FILE))
                if os.path.isdir(constants.ATTACHMENTS_DIR):
                    for root, _dirs, files in os.walk(constants.ATTACHMENTS_DIR):
                        for fname in files:
                            full_path = os.path.join(root, fname)
                            arcname = os.path.join(
                                "attachments",
                                os.path.relpath(full_path, constants.ATTACHMENTS_DIR)
                            )
                            zf.write(full_path, arcname=arcname)
            QMessageBox.information(self, "تم", "تم إنشاء النسخة الاحتياطية بنجاح.")
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء إنشاء النسخة الاحتياطية: {e}")