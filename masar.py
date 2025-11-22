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
from PyQt5 import QtCore
from PyQt5.QtCore import Qt, pyqtSignal
from weasyprint import HTML, CSS
import mimetypes
import shutil
import base64
import calendar
import time
from utils.pdf_bg_utils import process_bg_image




if __name__ == "__main__":
    init_db()
    app = QApplication(sys.argv)
    window = MasarMainWindow()
    window.showMaximized()
    sys.exit(app.exec_())
