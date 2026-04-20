# masar.py
from PyQt5.QtWidgets import QApplication, QDialog
from PyQt5.QtGui import QFontDatabase
from ui.main_window import MasarMainWindow
from ui.dialogs.license_dialog import LicenseDialog
from database.db_manager import init_db
from utils.license_core import LicenseManager
import sys
import os
import shutil
import platform
import ctypes

if __name__ == "__main__":
    init_db()
    app = QApplication(sys.argv)

    # Install Amiri font if not present on Windows
    if platform.system() == 'Windows':
        db = QFontDatabase()
        if 'Amiri' not in db.families():
            font_path = os.path.join(sys._MEIPASS if hasattr(sys, '_MEIPASS') else '.', 'assets', 'Amiri-Regular.ttf')
            dest = r'C:\Windows\Fonts\Amiri-Regular.ttf'
            if os.path.exists(font_path) and not os.path.exists(dest):
                try:
                    shutil.copy(font_path, dest)
                    ctypes.windll.gdi32.AddFontResourceW(dest)
                except Exception as e:
                    print(f"Failed to install Amiri font: {e}")

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
   pip install -r requirements.txt
   
2. لحزم التطبيق كملف exe:
   - ثبت pyinstaller: pip install pyinstaller
   - شغل الأمر:
   pyinstaller --onefile --windowed --icon "assets/icons/masar.ico" --add-data "assets;assets" --add-data "config.json;." --add-data "license.json;." --add-data "utils/pdf_bg_utils.py;." --add-data "attachments;attachments" --hidden-import=cryptography.fernet main.py

   سيظهر الملف التنفيذي في مجلد dist.

3. تأكد من نسخ مجلد attachments بجانب الملف التنفيذي.

"""

# Copyright 2025 Shehab.Magdy.Eladl
# Licensed under the Apache License, Version 2.0 (see LICENSE file)