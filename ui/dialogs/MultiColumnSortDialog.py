from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, 
    QPushButton, QListWidget, QListWidgetItem, QMessageBox, QWidget
)
from PyQt5.QtCore import Qt

class MultiColumnSortDialog(QDialog):
    def __init__(self, columns, parent=None):
        super().__init__(parent)
        self.setWindowTitle("ترتيب مخصص")
        self.resize(400, 300)
        self.columns = columns  # List of (index, name) tuples or just names
        self.sort_criteria = [] # List of (col_index, ascending_bool)

        layout = QVBoxLayout()
        
        # Selection area
        form_layout = QHBoxLayout()
        self.combo_columns = QComboBox()
        # Populate columns (assuming columns is list of strings)
        # If columns is a dict or list of objects, adapt here.
        # We will pass list of column names, index matches combo index
        self.combo_columns.addItems(self.columns)
        
        self.combo_order = QComboBox()
        self.combo_order.addItems(["تصاعدي", "تنازلي"])
        
        btn_add = QPushButton("إضافة مستوى")
        btn_add.clicked.connect(self.add_level)
        
        form_layout.addWidget(QLabel("العمود:"))
        form_layout.addWidget(self.combo_columns)
        form_layout.addWidget(self.combo_order)
        form_layout.addWidget(btn_add)
        
        layout.addLayout(form_layout)
        
        # List of levels
        self.levels_list = QListWidget()
        layout.addWidget(QLabel("مستويات الترتيب:"))
        layout.addWidget(self.levels_list)
        
        btn_remove = QPushButton("حذف المستوى المحدد")
        btn_remove.clicked.connect(self.remove_level)
        layout.addWidget(btn_remove)
        
        # Dialog buttons
        btns = QHBoxLayout()
        btn_ok = QPushButton("تطبيق")
        btn_ok.clicked.connect(self.accept)
        btn_cancel = QPushButton("إلغاء")
        btn_cancel.clicked.connect(self.reject)
        
        btns.addWidget(btn_ok)
        btns.addWidget(btn_cancel)
        layout.addLayout(btns)
        
        self.setLayout(layout)
        self.setLayoutDirection(Qt.RightToLeft)

    def add_level(self):
        col_idx = self.combo_columns.currentIndex()
        col_name = self.combo_columns.currentText()
        is_asc = (self.combo_order.currentIndex() == 0)
        
        # Check if column already in criteria
        for c, _ in self.sort_criteria:
            if c == col_idx:
                QMessageBox.warning(self, "تنبيه", "هذا العمود موجود بالفعل في القائمة")
                return

        self.sort_criteria.append((col_idx, is_asc))
        
        order_str = "تصاعدي" if is_asc else "تنازلي"
        item_text = f"{col_name} ({order_str})"
        self.levels_list.addItem(item_text)

    def remove_level(self):
        row = self.levels_list.currentRow()
        if row < 0:
            return
        self.levels_list.takeItem(row)
        self.sort_criteria.pop(row)

    def get_criteria(self):
        return self.sort_criteria
