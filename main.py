# masar.py
from PyQt5.QtWidgets import QApplication, QDialog
from ui.main_window import MasarMainWindow
from ui.dialogs.license_dialog import LicenseDialog
from database.db_manager import init_db
from utils.license_core import LicenseManager
import sys

if __name__ == "__main__":
    init_db()
    app = QApplication(sys.argv)

    # License check
    manager = LicenseManager()
    if not manager.validate_license():
        dialog = LicenseDialog(manager)
        if dialog.exec_() != QDialog.Accepted:
            sys.exit(0)  # Exit if activation failed or dialog closed

    window = MasarMainWindow()
    window.showMaximized()
    sys.exit(app.exec_())

 
"""
تعليمات البناء (Build Instructions):

1. تأكد من تثبيت المتطلبات:
   pip install pyqt5 pillow weasyprint
   
2. لحزم التطبيق كملف exe:
   - ثبت pyinstaller: pip install pyinstaller
   - شغل الأمر:
   pyinstaller --onefile --windowed --icon "assets/icons/masar.ico" --add-data "assets/masar-bg.png;." --add-data "assets/Amiri-Regular.ttf;." --add-data "config.json;." --add-data "license.json;." --add-data "utils/pdf_bg_utils.py;." --add-data "attachments;attachments" main.py

   سيظهر الملف التنفيذي في مجلد dist.

3. تأكد من نسخ مجلد attachments بجانب الملف التنفيذي.

"""

# Copyright 2025 Shehab.Magdy.Eladl
# Licensed under the Apache License, Version 2.0 (see LICENSE file)