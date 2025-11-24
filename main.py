# masar.py
from PyQt5.QtWidgets import QApplication
from ui.main_window import MasarMainWindow
from database.db_manager import init_db
import sys

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