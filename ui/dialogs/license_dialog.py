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
        self.hw_edit.setText(self.manager.encrypt_A(hw_id))
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