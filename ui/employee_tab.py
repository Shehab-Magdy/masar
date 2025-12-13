# ui/employee_tab.py
import sys
import os
import datetime
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QPushButton,
    QTableWidget, QTableWidgetItem, QLineEdit, QHBoxLayout, QFileDialog, QListWidget,
    QMessageBox, QSizePolicy, QGridLayout, QDialog
)
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt
from weasyprint import HTML, CSS
from ui.dialogs.MultiColumnSortDialog import MultiColumnSortDialog
import mimetypes
import shutil
import base64
from utils.arabic_normalizer import normalize_arabic
from utils.constants import AR_LABELS, EMPLOYEE_FIELDS, cfg_path, bg_path
from utils.file_utils import get_employee_folder
from utils.pdf_bg_utils import process_bg_image

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
        self.table.setLayoutDirection(Qt.LayoutDirection.RightToLeft)

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
        
        self.btn_search_advanced = QPushButton("بحث بالمواصفات")
        self.btn_search_advanced.clicked.connect(self.search_advanced)
        btns_layout.addWidget(self.btn_search_advanced)

        self.btn_sort = QPushButton("ترتيب مخصص")
        self.btn_sort.clicked.connect(self.custom_sort_action)
        btns_layout.addWidget(self.btn_sort)

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
        Also resets the search field and reloads the full employee list (clearing filters/sorts).
        :return: None
        :rtype: NoneType
        """
        self.search_field.clear()
        for f in EMPLOYEE_FIELDS:
            self.form_fields[f].clear()
        self.attach_list.clear()
        self.attachments = []
        self.photo_path = None
        self.photo_label.clear()
        self.selected_emp_id = None
        self.load_employees()

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

    def search_advanced(self):
        """
        Searches for employees based on values entered in the form fields.
        Combines non-empty fields with AND logic using partial matching (LIKE).
        """
        self.table.setRowCount(0)
        c = self.conn.cursor()
        
        conditions = []
        params = []
        
        for f in EMPLOYEE_FIELDS:
            val = self.form_fields[f].text().strip()
            if val:
                # Normalize logic for Arabic text if needed
                norm_val = normalize_arabic(val)
                conditions.append(f"{f} LIKE ?")
                params.append(f"%{norm_val}%")
        
        if not conditions:
             # If no criteria entered, maybe load all
             self.load_employees()
             return

        query = f"SELECT id, {', '.join(EMPLOYEE_FIELDS)} FROM employee WHERE " + " AND ".join(conditions)
        
        c.execute(query, tuple(params))
        rows = c.fetchall()
        
        if not rows:
            QMessageBox.information(self, "بحث", "لا توجد نتائج مطابقة")
            return

        for row in rows:
            row_idx = self.table.rowCount()
            self.table.insertRow(row_idx)
            for col_idx, val in enumerate(row[1:]):
                self.table.setItem(row_idx, col_idx, QTableWidgetItem(str(val)))
            self.table.setVerticalHeaderItem(row_idx, QTableWidgetItem(str(row[0])))



    def custom_sort_action(self):
        """
        Opens the MultiColumnSortDialog and sorts the current table contents locally.
        """
        # Get column names from headers
        headers = []
        for i in range(self.table.columnCount()):
            item = self.table.horizontalHeaderItem(i)
            headers.append(item.text() if item else str(i))
            
        dlg = MultiColumnSortDialog(headers, self)
        if dlg.exec() == QDialog.Accepted:
            criteria = dlg.get_criteria() # list of (col_idx, is_asc)
            if not criteria:
                return
                
            # Scrape data from table
            rows_data = []
            row_count = self.table.rowCount()
            col_count = self.table.columnCount()
            
            for r in range(row_count):
                row_items = []
                # Also capture vertical header (ID)
                vh_item = self.table.verticalHeaderItem(r)
                row_id = vh_item.text() if vh_item else ""
                
                for c in range(col_count):
                    item = self.table.item(r, c)
                    txt = item.text() if item else ""
                    row_items.append(txt)
                
                rows_data.append((row_id, row_items))
                
            # Sort data
            # sort is stable, so we can sort in reverse order of criteria
            # to achieve multi-level sort effect.
            # However, standard Python sort with key tuple is easier if all same direction.
            # But here directions might differ.
            # Stable sort approach: apply sorts from last criterion to first.
            
            for col_idx, is_asc in reversed(criteria):
                # Try to sort numerically if possible, else string
                def sort_key(row_tuple):
                    val = row_tuple[1][col_idx]
                    # basic heuristic for number
                    # Return tuple (type_priority, value) to safely compare int vs str
                    # type_priority: 0 for int, 1 for str
                    # This ensures all ints are compared together and all strs together
                    if val.isdigit():
                        return (0, int(val))
                    return (1, val)
                
                rows_data.sort(key=sort_key, reverse=not is_asc)
                
            # Re-populate table
            self.table.setRowCount(0) # clear rows
            for r_idx, (r_id, r_items) in enumerate(rows_data):
                self.table.insertRow(r_idx)
                for c_idx, val in enumerate(r_items):
                    self.table.setItem(r_idx, c_idx, QTableWidgetItem(val))
                self.table.setVerticalHeaderItem(r_idx, QTableWidgetItem(r_id))

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
        headers2 = [AR_LABELS[f] for f in fields2]

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

        if os.path.isfile(cfg_path):
            try:
                import json
                with open(cfg_path, 'r', encoding='utf-8') as f:
                    cfg = json.load(f)
                first_line_header = cfg.get('firstLineHeader', "")
                second_line_header = cfg.get('secondLineHeader', "")
                font_size = cfg.get('font-size', 11)
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
                    src: url('assets/Amiri-Regular.ttf') format('truetype');
                }}
                body {{
                    direction: rtl;
                    font-family: 'Amiri', 'Cairo', 'Tahoma', sans-serif;
                    font-size: {font_size}px;
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
                /* light rows transparent so the PDF background shows through */
                tr.zebra1 {{ background-color: transparent; }}
                tr.zebra2 {{ background-color: #f2f2f2; }}
                @page {{
                    size: A4 landscape;
                    margin: 1cm 1cm 2cm 1cm;
                    @bottom-center {{
                        content: "الصفحة " counter(page) " من " counter(pages);
                        font-family: 'Amiri', 'Cairo', 'Tahoma', sans-serif;
                        font-size: 12px;
                        color: #444;
                    }}
                }}
            </style>
        </head>
        <body>
            <div style="text-align:right; margin-bottom: 8px;">
                <div style="font-size:13px; color:#1976d2;">{first_line_header}</div>
                <div style="font-size:13px; color:#1976d2;">{second_line_header}</div>
            </div>
            <h2 style="font-size:15px; text-align:center;">بيانات الموظفين المدنيين في الورش الرئيسية للطائرات</h2>
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
                    margin: 1cm 0.5cm 1.5cm 0.5cm;
                    @top-right {{
                        content: '""" + first_line_header + """\\A""" + second_line_header + """';
                        font-size: 15px;
                        color: #1976d2;
                        text-align: right;
                        white-space: pre;
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
                            content: '""" + first_line_header + """\\A""" + second_line_header + """';
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
