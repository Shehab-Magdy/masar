# ui/main_window.py
import sqlite3
from PyQt5.QtWidgets import QMainWindow, QTabWidget
from ui.dashboard_tab import DashboardTab
from ui.employee_tab import EmployeeTab
from ui.correspondence_tab import CorrespondenceTab
from database.db_manager import init_db
import os
from PyQt5.QtGui import QIcon
from utils.constants import DB_FILE
from utils.file_utils import resource_path

class MasarMainWindow(QMainWindow):
    def __init__(self):
        """
        Initializes the main window of the application.

        Sets the window title, geometry, and icon.
        Connects to the database and sets up the tab widget with the three main tabs:
        "لوحة التحكم", "الموظفين", and "التقارير".
        """
        super().__init__()
        self.setWindowTitle("مسار - منظومة العاملين المدنيين بالورش")
        self.setGeometry(100, 100, 1100, 700)
        # Set window icon
        self.setWindowIcon(QIcon(resource_path("assets/masar.ico")))
        self.conn = sqlite3.connect(DB_FILE)
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)
        self.tabs.addTab(DashboardTab(self.conn), "الإحصائيات")
        self.tabs.addTab(EmployeeTab(self.conn), "الموظفين")
        # Correspondence tab (المراسلات)
        try:
            self.tabs.addTab(CorrespondenceTab(self.conn), "المراسلات")
        except NameError:
            # Class may be defined later in file; adding tab will work after import
            pass
