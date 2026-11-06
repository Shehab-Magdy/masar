from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QMessageBox, QApplication, QTextEdit
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QClipboard
from utils.license_core import LicenseManager
import sys
import os

class LicenseDialog(QDialog):
    def __init__(self, manager=None):
        super().__init__()
        self.manager = manager or LicenseManager()
        self.setWindowTitle("تفعيل الترخيص - License Activation")
        self.setModal(True)
        self.setFixedSize(500, 300)
        self.setWindowFlags(Qt.WindowCloseButtonHint | Qt.WindowTitleHint)

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Hardware ID Section
        hw_label = QLabel("معرف الجهاز (Hardware ID):")
        layout.addWidget(hw_label)

        hw_layout = QHBoxLayout()
        self.hw_edit = QLineEdit()
        self.hw_edit.setReadOnly(True)
        hw_id = self.manager.get_hardware_id()
        self.hw_edit.setText(hw_id)
        hw_layout.addWidget(self.hw_edit)

        copy_btn = QPushButton("نسخ")
        copy_btn.clicked.connect(self.copy_hw_id)
        hw_layout.addWidget(copy_btn)

        layout.addLayout(hw_layout)

        # Instructions
        instr_label = QLabel("أرسل معرف الجهاز إلى الإدارة للحصول على كود التفعيل")
        instr_label.setStyleSheet("color: gray; font-size: 10px;")
        layout.addWidget(instr_label)

        # Activation Code Section
        code_label = QLabel("كود التفعيل (Activation Code):")
        layout.addWidget(code_label)

        self.code_edit = QLineEdit()
        self.code_edit.setPlaceholderText("أدخل كود التفعيل هنا")
        self.code_edit.textChanged.connect(self._format_code_input)
        layout.addWidget(self.code_edit)

        # Buttons
        btn_layout = QHBoxLayout()

        activate_btn = QPushButton("تفعيل")
        activate_btn.clicked.connect(self.activate_license)
        btn_layout.addWidget(activate_btn)

        exit_btn = QPushButton("خروج")
        exit_btn.clicked.connect(self.close_app)
        btn_layout.addWidget(exit_btn)

        layout.addLayout(btn_layout)

        self.setLayout(layout)

    def _format_code_input(self, text):
        """Auto-uppercase and auto-insert dashes every 4 chars as the user types/pastes"""
        cursor = self.code_edit.cursorPosition()
        alnum_before = sum(1 for ch in text[:cursor] if ch.isalnum())
        normalized = ''.join(ch for ch in text.upper() if ch.isalnum())[:16]

        grouped_chars = []
        new_cursor = 0
        for i, ch in enumerate(normalized):
            if i > 0 and i % 4 == 0:
                grouped_chars.append('-')
            grouped_chars.append(ch)
            if i + 1 == alnum_before:
                new_cursor = len(grouped_chars)
        if alnum_before >= len(normalized):
            new_cursor = len(grouped_chars)

        grouped = ''.join(grouped_chars)
        if grouped != text:
            self.code_edit.blockSignals(True)
            self.code_edit.setText(grouped)
            self.code_edit.setCursorPosition(new_cursor)
            self.code_edit.blockSignals(False)

    def copy_hw_id(self):
        clipboard = QApplication.clipboard()
        clipboard.setText(self.hw_edit.text())
        QMessageBox.information(self, "تم النسخ", "تم نسخ معرف الجهاز إلى الحافظة!")

    def activate_license(self):
        code = self.code_edit.text().strip()
        if not code:
            QMessageBox.warning(self, "خطأ", "يرجى إدخال كود التفعيل")
            return

        try:
            success = self.manager.activate_license(code)
            if success:
                QMessageBox.information(self, "نجح التفعيل", "تم التفعيل بنجاح. سيتم إعادة تشغيل التطبيق...")
                self.accept()  # Close dialog and continue
            else:
                QMessageBox.warning(self, "فشل التفعيل", "كود التفعيل غير صحيح أو لا يتطابق مع الجهاز")
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء التفعيل: {str(e)}")

    def close_app(self):
        sys.exit(0)