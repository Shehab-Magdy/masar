from PyQt5.QtWidgets import QLabel
from PyQt5.QtCore import pyqtSignal

class ClickableLabel(QLabel):
    """A QLabel that emits a clicked signal when pressed."""
    clicked = pyqtSignal()

    def mousePressEvent(self, event):
        if event.button() == 1:
            self.clicked.emit()
        super().mousePressEvent(event)
