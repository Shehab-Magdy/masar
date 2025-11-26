from PyQt5.QtWidgets import QCalendarWidget
from PyQt5.QtGui import QTextCharFormat, QColor, QFont
from PyQt5.QtCore import Qt, QDate

class CustomWeekendCalendar(QCalendarWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
        self.setFirstDayOfWeek(Qt.DayOfWeek.Saturday)
        self.highlight_weekends()
        

    def highlight_weekends(self):
        # --- Weekend text color (Friday + Saturday) ---
        weekend_format = QTextCharFormat()
        weekend_format.setForeground(QColor("#D9534F"))  # soft red
        for day in [Qt.DayOfWeek.Friday, Qt.DayOfWeek.Saturday]:
            self.setWeekdayTextFormat(day, weekend_format)

        # --- Weekday text color ---
        weekday_format = QTextCharFormat()
        weekday_format.setForeground(QColor("#1D4E89"))  # Masar navy
        for day in [
            Qt.DayOfWeek.Sunday, Qt.DayOfWeek.Monday, Qt.DayOfWeek.Tuesday,
            Qt.DayOfWeek.Wednesday, Qt.DayOfWeek.Thursday
        ]:
            self.setWeekdayTextFormat(day, weekday_format)

        # --- Highlight today's date ---
        today_format = QTextCharFormat()
        today_format.setForeground(QColor("#1D4E89"))
        today_format.setBackground(QColor("#DCE7FF"))
        today_format.setFontWeight(QFont.Bold)
        self.setDateTextFormat(QDate.currentDate(), today_format)

        # --- Apply stylesheet for hover, rounded cells, etc. ---
        self.setStyleSheet("""
            QCalendarWidget QWidget {
                background-color: #F5F7FA;
                color: #1D4E89;
            }

            QCalendarWidget QAbstractItemView:enabled {
                background-color: #FFFFFF;
                color: #1D4E89;
                selection-background-color: #3E7BFA;
                selection-color: #FFFFFF;
                border-radius: 6px;
            }

            /* Hover Effect */
            QCalendarWidget QAbstractItemView:item:hover {
                background-color: #E3ECFF;
                border-radius: 6px;
            }

            /* Navigation buttons */
            QCalendarWidget QToolButton {
                background-color: #E9EEF5;
                color: #1D4E89;
                border-radius: 6px;
                padding: 4px;
            }

            QCalendarWidget QToolButton:hover {
                background-color: #DCE4EF;
            }

            QCalendarWidget QMenu {
                background-color: #FFFFFF;
                color: #1D4E89;
            }
        """)