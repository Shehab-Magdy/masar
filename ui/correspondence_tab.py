import sys
import os
import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QLabel, QPushButton, QHeaderView,
    QTableWidget, QTableWidgetItem, QLineEdit, QHBoxLayout, QFileDialog, QListWidget,
    QMessageBox, QTextEdit, QGridLayout, QDateEdit, QListWidgetItem, QComboBox, QSizePolicy
)
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt
from weasyprint import HTML, CSS
import base64
import time
from ui.dialogs.ClickableLabel import ClickableLabel
from utils.arabic_normalizer import normalize_arabic
from utils.constants import ATTACHMENTS_DIR, bg_path, cfg_path
from utils.pdf_bg_utils import process_bg_image
from reports.correspondence_export import export_correspondence_pdf
from ui.dialogs.CustomWeekendCalendar import CustomWeekendCalendar


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
        self.btn_export_results = QPushButton("تصدير النتائج")
        self.btn_export_results.clicked.connect(self.export_visible_results)
        search_layout.addWidget(self.btn_export_results)
        
        self.btn_show_all = QPushButton("عرض الكل")
        self.btn_show_all.clicked.connect(self.show_all_entries)
        search_layout.addWidget(self.btn_show_all)

        self.btn_search = QPushButton("بحث")
        self.btn_search.clicked.connect(self.search_entries)
        search_layout.addWidget(self.btn_search)
        
        self.search_subject = QLineEdit()
        self.search_subject.setPlaceholderText("بحث في الموضوع...")
        search_layout.addWidget(self.search_subject)
        search_layout.addWidget(QLabel("الموضوع:"))

        self.search_fax = QLineEdit()
        self.search_fax.setPlaceholderText("بحث برقم الفاكس...")
        search_layout.addWidget(self.search_fax)
        search_layout.addWidget(QLabel("رقم الفاكس:"))
        
        self.search_to = QDateEdit()
        self.search_to.setCalendarPopup(True)
        self.search_to.setDisplayFormat("yyyy-MM-dd")
        self.search_to.setDate(datetime.date.today())
        search_layout.addWidget(self.search_to)
        search_layout.addWidget(QLabel("إلى التاريخ:"))

        self.search_from = QDateEdit()
        self.search_from.setCalendarPopup(True)
        self.search_from.setDisplayFormat("yyyy-MM-dd")
        self.search_from.setDate(datetime.date.today())
        search_layout.addWidget(self.search_from)
        search_layout.addWidget(QLabel("من التاريخ:"))

        main_layout.addLayout(search_layout)

        # Main content layout: table and form
        content_layout = QHBoxLayout()

        # Table of correspondence
        self.table = QTableWidget()
        headers = ["مسلسل", "رقم الفاكس", "نوع المراسلة", "التاريخ", "صادر من", "وارد الى", "الموضوع", "الملاحظات"]
        self.table.setColumnCount(8)
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setSelectionBehavior(self.table.SelectRows)
        self.table.cellClicked.connect(self.on_row_select)
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet("QTableWidget {alternate-background-color: #f2f2f2; background-color: transparent;}")
        self.table.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        header = self.table.horizontalHeader()
        if header is not None:
            header.setSectionResizeMode(QHeaderView.Stretch)
        content_layout.addWidget(self.table, 2)

        # --- Form fields dictionary for easy access ---
        self.form_fields = {
            "fax_number": QLineEdit(),
            "from_person": QLineEdit(),
            "fax_date": QDateEdit(),
            "subject": QLineEdit(),
            "correspondence_type": QComboBox(),
            "to_person": QLineEdit(),
            "notes": QTextEdit(),
        }

        self.form_fields["correspondence_type"].addItems(['اختر', 'صادر', 'وارد'])
        self.form_fields["fax_date"].setCalendarPopup(True)
        self.form_fields["fax_date"].setDisplayFormat("yyyy-MM-dd")
        self.form_fields["fax_date"].setDate(datetime.date.today())


        # --- Labels in Arabic ---
        AR_LABELS = {
            "fax_number": "رقم الفاكس",
            "correspondence_type": "نوع المراسلة",
            "fax_date": "التاريخ",
            "from_person": "صادر من",
            "to_person": "وارد إلى",
            "subject": "الموضوع",
            "notes": "الملاحظات",
        }

        # --- Grid layout ---
        grid_layout = QGridLayout()
        grid_layout.setAlignment(Qt.AlignmentFlag.AlignRight)
        grid_layout.setHorizontalSpacing(12)
        grid_layout.setVerticalSpacing(8)

        # Row 0: رقم الفاكس | نوع المراسلة
        grid_layout.addWidget(QLabel("رقم الفاكس:"), 0, 0)
        grid_layout.addWidget(self.form_fields["fax_number"], 0, 1)
        grid_layout.addWidget(QLabel("نوع المراسلة:"), 0, 2)
        grid_layout.addWidget(self.form_fields["correspondence_type"], 0, 3)

        # Row 1: التاريخ (alone, spanning columns 1-3)
        grid_layout.addWidget(QLabel("التاريخ:"), 1, 0)
        grid_layout.addWidget(self.form_fields["fax_date"], 1, 1, 1, 3)

        # Row 2: صادر من | وارد إلى
        grid_layout.addWidget(QLabel("صادر من:"), 2, 0)
        grid_layout.addWidget(self.form_fields["from_person"], 2, 1)
        grid_layout.addWidget(QLabel("وارد إلى:"), 2, 2)
        grid_layout.addWidget(self.form_fields["to_person"], 2, 3)

        # Row 3: الموضوع (alone, spanning columns 1-3)
        grid_layout.addWidget(QLabel("الموضوع:"), 3, 0)
        grid_layout.addWidget(self.form_fields["subject"], 3, 1, 1, 3)

        # Row 4: الملاحظات (alone, spanning columns 1-3)
        grid_layout.addWidget(QLabel("الملاحظات:"), 4, 0)
        grid_layout.addWidget(self.form_fields["notes"], 4, 1, 1, 3)

        row_offset = 5

        # Attachments and image controls (row_offset is the next available row)
        attachments_layout = QVBoxLayout()
        attachments_layout.addWidget(QLabel("رفع صور الفاكس:"))
        btn_img = QPushButton("اختر صور")
        btn_img.clicked.connect(self.browse_image)
        attachments_layout.addWidget(btn_img)
        btn_remove_img = QPushButton("حذف الصورة المحددة")
        btn_remove_img.clicked.connect(self.remove_selected_image)
        attachments_layout.addWidget(btn_remove_img)
        attachments_widget = QWidget()
        attachments_widget.setLayout(attachments_layout)

        self.fax_images_list = QListWidget()
        self.fax_images_list.setFixedHeight(90)

        # Horizontal layout: attachments_widget (right), stretch, fax_images_list (left)
        attachments_row_layout = QHBoxLayout()
        attachments_row_layout.addWidget(attachments_widget, alignment=Qt.AlignmentFlag.AlignTop)
        attachments_row_layout.addStretch(1)
        attachments_row_layout.addWidget(self.fax_images_list)

        # Place the horizontal layout in the grid (spanning all columns)
        grid_layout.addLayout(attachments_row_layout, row_offset, 0, 1, 4)
        self.fax_images_list.itemClicked.connect(self.on_image_item_clicked)
        row_offset += 1
        
        # Thumbnail preview section (row_offset is the next available row)
        thumbnail_layout = QVBoxLayout()
        thumbnail_layout.addWidget(QLabel("معاينة الصورة:"))
        # If you want to add a button under the label, do it here (optional)
        btn_open = QPushButton("فتح الصورة")
        btn_open.clicked.connect(self.open_current_thumbnail)
        thumbnail_layout.addWidget(btn_open)
        thumbnail_widget = QWidget()
        thumbnail_widget.setLayout(thumbnail_layout)

        # Thumbnail image widget
        self.thumbnail = ClickableLabel()
        self.thumbnail.setMinimumHeight(150)
        self.thumbnail.setMinimumWidth(200)
        self.thumbnail.setStyleSheet("border: 1px solid #ccc; background: #fff;")
        self.thumbnail.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.thumbnail.clicked.connect(self.open_current_thumbnail)

        # Horizontal layout: thumbnail_widget (right), stretch, thumbnail image (left)
        thumbnail_row_layout = QHBoxLayout()
        thumbnail_row_layout.addWidget(thumbnail_widget, alignment=Qt.AlignmentFlag.AlignTop)
        thumbnail_row_layout.addStretch(1)
        thumbnail_row_layout.addWidget(self.thumbnail)

        # Place the horizontal layout in the grid (spanning all columns)
        grid_layout.addLayout(thumbnail_row_layout, row_offset, 0, 1, 4)
        row_offset += 1

        # Action buttons (span all columns)
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
        grid_layout.addLayout(btns, row_offset, 0, 1, 4)

        form_widget = QWidget()
        form_widget.setLayout(grid_layout)
        form_widget.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        content_layout.addWidget(form_widget, 1)

        main_layout.addLayout(content_layout)

        self.setLayout(main_layout)
        self.load_entries()

        # For each QDateEdit field:
        calendar = CustomWeekendCalendar()
        self.form_fields["fax_date"].setCalendarWidget(calendar)

        calendar_from = CustomWeekendCalendar()
        self.search_from.setCalendarWidget(calendar_from)

        calendar_to = CustomWeekendCalendar()
        self.search_to.setCalendarWidget(calendar_to)

    def browse_image(self):
        files, _ = QFileDialog.getOpenFileNames(self, "اختر صور الفاكس", "", "Images (*.png *.jpg *.jpeg)")
        if not files:
            return
        for f in files:
            if f:
                self.temp_images.append(f)
                item = QListWidgetItem(os.path.basename(f))
                item.setData(Qt.ItemDataRole.UserRole, ("temp", f))
                self.fax_images_list.addItem(item)
        # show thumbnail of the last selected image
        if files:
            self._show_thumbnail_from_path(files[-1])

    def clear_form(self):
        self.selected_id = None
        self.form_fields["fax_number"].clear()
        # Reset fax_date to today
        self.form_fields["fax_date"].setDate(datetime.date.today())
        # Reset correspondence_type to 'اختر'
        self.form_fields["correspondence_type"].setCurrentText('اختر')
        self.form_fields["from_person"].clear()
        self.form_fields["to_person"].clear()
        self.form_fields["subject"].clear()
        self.form_fields["notes"].clear()
        self.temp_images = []
        self.current_attachments = []
        self.fax_images_list.clear()
        # Clear thumbnail content
        self.thumbnail.clear()
        self.current_thumbnail_path = None

    def load_entries(self, where_clause: str = "", params: tuple = ()): 
        c = self.conn.cursor()
        # Add correspondence_type to the SELECT and table columns
        q = "SELECT id, fax_number, correspondence_type, fax_date, from_person, to_person, subject, notes, created_at FROM correspondence"
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
            self.table.setItem(idx, 1, QTableWidgetItem(str(row[1] or "")))  # fax_number
            self.table.setItem(idx, 2, QTableWidgetItem(str(row[2] or "")))  # correspondence_type
            self.table.setItem(idx, 3, QTableWidgetItem(str(row[3] or "")))  # fax_date
            self.table.setItem(idx, 4, QTableWidgetItem(str(row[4] or "")))  # from_person
            self.table.setItem(idx, 5, QTableWidgetItem(str(row[5] or "")))  # to_person
            self.table.setItem(idx, 6, QTableWidgetItem(str(row[6] or "")))  # subject
            self.table.setItem(idx, 7, QTableWidgetItem(str(row[7] or "")))  # notes
            # show count of images for the correspondence
            c2 = self.conn.cursor()
            c2.execute("SELECT COUNT(*) FROM correspondence_attachment WHERE correspondence_id=?", (rid,))
            cnt = c2.fetchone()[0]
            # If you want to show image count in a new column, add another column and set it here
            # self.table.setItem(idx, 8, QTableWidgetItem(str(cnt)))
            self.table.setVerticalHeaderItem(idx, QTableWidgetItem(str(row[0])))


    def on_row_select(self, row, col):
        vh = self.table.verticalHeaderItem(row)
        if vh is None:
            return
        rid = vh.text()
        c = self.conn.cursor()
        c.execute("SELECT id, fax_number, fax_date, from_person, to_person, subject, notes, correspondence_type FROM correspondence WHERE id=?", (rid,))
        r = c.fetchone()
        if not r:
            return
        self.selected_id = r[0]
        self.form_fields["fax_number"].setText(r[1] or "")
        try:
            if r[2]:
                dt = datetime.datetime.strptime(r[2], "%Y-%m-%d").date()
                self.form_fields["fax_date"].setDate(dt)
        except Exception:
            pass
        self.form_fields["from_person"].setText(r[3] or "")
        self.form_fields["to_person"].setText(r[4] or "")
        self.form_fields["subject"].setText(r[5] or "")
        self.form_fields["notes"].setPlainText(r[6] or "")
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
            item.setData(Qt.ItemDataRole.UserRole, ("stored", att_id, fpath))
            self.fax_images_list.addItem(item)
        # show thumbnail of first attachment if available
        if self.current_attachments:
            first_path = self.current_attachments[0][2]
            if first_path:
                self._show_thumbnail_from_path(first_path)
        self.form_fields["correspondence_type"].setCurrentText(r[7] if len(r) > 7 and r[7] else 'صادر')

    def _validate_form(self):
        fax_no = self.form_fields["fax_number"].text().strip()
        subj = self.form_fields["subject"].text().strip()
        corr_type = self.form_fields["correspondence_type"].currentText()
        if not fax_no:
            return False, "رقم الفاكس مطلوب"
        if corr_type == 'اختر':
            return False, "من فضلك اختر نوع المراسلة (صادر أو وارد)."
        if not subj:
            return False, "الموضوع مطلوب"
        # date validation
        d = self.form_fields["fax_date"].date().toPyDate()
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
                item.setData(Qt.ItemDataRole.UserRole, ("stored", att_id, fpath))
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
        fax_no = self.form_fields["fax_number"].text().strip()
        # Check for duplicate fax number
        c = self.conn.cursor()
        c.execute("SELECT COUNT(*) FROM correspondence WHERE fax_number = ?", (fax_no,))
        if c.fetchone()[0] > 0:
            QMessageBox.warning(self, "خطأ", "رقم الفاكس موجود بالفعل ولا يمكن تكراره.")
            return
        fax_date = self.form_fields["fax_date"].date().toString("yyyy-MM-dd")
        from_p = normalize_arabic(self.form_fields["from_person"].text().strip())
        to_p = normalize_arabic(self.form_fields["to_person"].text().strip())
        subj = normalize_arabic(self.form_fields["subject"].text().strip())
        notes = self.form_fields["notes"].toPlainText().strip()
        created_at = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        # insert correspondence record first (image attachments will be stored in correspondence_attachment)
        c.execute("INSERT INTO correspondence (fax_number, correspondence_type, fax_date, from_person, to_person, subject, notes, image_path, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)", (fax_no, self.form_fields["correspondence_type"].currentText(), fax_date, from_p, to_p, subj, notes, '', created_at))
        
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
        fax_no = self.form_fields["fax_number"].text().strip()
        fax_date = self.form_fields["fax_date"].date().toString("yyyy-MM-dd")
        from_p = normalize_arabic(self.form_fields["from_person"].text().strip())
        to_p = normalize_arabic(self.form_fields["to_person"].text().strip())
        subj = normalize_arabic(self.form_fields["subject"].text().strip())
        notes = self.form_fields["notes"].toPlainText().strip()
        c = self.conn.cursor()
        c.execute(
    "UPDATE correspondence SET fax_number=?, correspondence_type=?, fax_date=?, from_person=?, to_person=?, subject=?, notes=? WHERE id=?",
    (fax_no, self.form_fields["correspondence_type"].currentText(), fax_date, from_p, to_p, subj, notes, self.selected_id)
)
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
        
        fax = self.search_fax.text().strip()
        if fax:
            clauses.append("fax_number LIKE ?")
            params.append(f"%{fax}%")
        where = " AND ".join(clauses)
        self.load_entries(where, tuple(params))

    def remove_selected_image(self):
        """Remove selected image from the list. If it's a stored attachment, delete from DB and disk."""
        item = self.fax_images_list.currentItem()
        if not item:
            return
        data = item.data(Qt.ItemDataRole.UserRole)
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
                d = it.data(Qt.ItemDataRole.UserRole)
                if d and d[0] == "stored" and d[1] == att_id:
                    self.fax_images_list.takeItem(i)
                    break
        # clear thumbnail if no items left
        if self.fax_images_list.count() == 0:
            self.thumbnail.clear()

    def on_image_item_clicked(self, item):
        if item is None:
            return
        data = item.data(Qt.ItemDataRole.UserRole)
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
            scaled = pix.scaled(self.thumbnail.width(), self.thumbnail.height(), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
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

    def show_all_entries(self):
        # Clear search fields
        self.search_subject.clear()
        self.search_fax.clear()
        self.search_from.setDate(datetime.date.today())
        self.search_to.setDate(datetime.date.today())
        # Reload all entries
        self.load_entries()
    
    def export_visible_results(self):
        # Get IDs of visible rows from the table's vertical header
        ids = []
        for row in range(self.table.rowCount()):
            vh = self.table.verticalHeaderItem(row)
            if vh:
                ids.append(vh.text())
        if not ids:
            QMessageBox.warning(self, "تنبيه", "لا توجد نتائج للتصدير")
            return
        now = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        file_path, _ = QFileDialog.getSaveFileName(self, "حفظ المراسلات كـ PDF", f"correspondence_{now}.pdf", "PDF Files (*.pdf)")
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

        # build html
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
                font-size: 11px; 
                {'background: url("'+bg_url+'") no-repeat center center; background-size: contain;' if bg_url else ''} }}
            table {{ 
                border-collapse: collapse; 
                width: 100%; 
                table-layout: fixed; 
            }}
            th, td {{ 
                border: 1px solid #888; 
                padding: 6px; 
                vertical-align: top; 
                text-align: right; 
                word-break: break-word;
                white-space: pre-wrap; 
            }}
            th {{ 
                background: #b3d1f7; 
            }}
            tr:nth-child(odd) {{ 
                background-color: transparent; 
            }}
            tr:nth-child(even) {{ 
                background-color: #f2f2f2; 
            }}
            @page {{ 
                size: A4 portrait; 
                margin: 1cm 0.5cm 1.5cm 0.5cm;
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
