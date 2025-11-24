from PyQt5.QtWidgets import QCalendarWidget
from PyQt5.QtGui import QTextCharFormat, QColor
from PyQt5.QtCore import Qt

class CustomWeekendCalendar(QCalendarWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
        self.setFirstDayOfWeek(Qt.DayOfWeek.Saturday)
        self.highlight_weekends()
        

    def highlight_weekends(self):
        # Example: Friday (5) and Saturday (6) as weekends
        weekend_format = QTextCharFormat()
        # weekend_format.setBackground(QColor("#ffeaea"))
        weekend_format.setForeground(QColor("#d32f2f"))
        for day in [Qt.DayOfWeek.Friday, Qt.DayOfWeek.Saturday]:
            self.setWeekdayTextFormat(day, weekend_format)
        
        weekday_format = QTextCharFormat()
        weekday_format.setForeground(QColor("#ffeaea"))
        for day in [Qt.DayOfWeek.Sunday, Qt.DayOfWeek.Monday, Qt.DayOfWeek.Tuesday, Qt.DayOfWeek.Wednesday, Qt.DayOfWeek.Thursday]:
            self.setWeekdayTextFormat(day, weekday_format)
        
        self.setStyleSheet("""
            QCalendarWidget QWidget {
                background-color: #2c2c2c;
                color: #ffeaea;
            }
            QCalendarWidget QAbstractItemView:enabled {
                background-color: #424242;
                color: #ffeaea;
                selection-background-color: #616161;
                selection-color: #ffeaea;
            }
            QCalendarWidget QToolButton {
                background-color: #616161;
                color: #ffeaea;
            }
            QCalendarWidget QMenu {
                background-color: #424242;
                color: #ffeaea;
            }
        """)