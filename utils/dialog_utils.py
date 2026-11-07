# utils/dialog_utils.py
"""
Centers every popup (QDialog subclasses, including QMessageBox and
QFileDialog) on the screen when it's shown, instead of leaving it wherever
Qt/the window manager would otherwise place it - most noticeably the
license dialog, which appears before the main window exists.
"""
from PyQt5.QtCore import QEvent, QObject
from PyQt5.QtWidgets import QApplication, QDialog


class DialogCenterFilter(QObject):
    def eventFilter(self, obj, event):
        if event.type() == QEvent.Show and isinstance(obj, QDialog):
            screen = QApplication.primaryScreen()
            if screen:
                geo = screen.availableGeometry()
                frame = obj.frameGeometry()
                frame.moveCenter(geo.center())
                obj.move(frame.topLeft())
        return super().eventFilter(obj, event)


def install_dialog_centering(app):
    """Call once on the QApplication instance at startup."""
    center_filter = DialogCenterFilter(app)
    app.installEventFilter(center_filter)
    app._dialog_center_filter = center_filter  # keep a reference alive
    return center_filter
