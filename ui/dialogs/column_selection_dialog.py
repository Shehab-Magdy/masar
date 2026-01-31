from PyQt5.QtWidgets import QDialog, QVBoxLayout, QCheckBox, QDialogButtonBox
from utils.constants import EMPLOYEE_FIELDS, AR_LABELS


class ColumnSelectionDialog(QDialog):
    """Dialog to select which employee columns to export.

    Usage:
        dlg = ColumnSelectionDialog(parent)
        if dlg.exec_() == QDialog.Accepted:
            selected = dlg.selected_fields()
    """
    def __init__(self, parent=None, fields=EMPLOYEE_FIELDS):
        super().__init__(parent)
        self.setWindowTitle("اختيار الأعمدة للتصدير")
        self.check_boxes = {}
        layout = QVBoxLayout()
        for f in fields:
            cb = QCheckBox(AR_LABELS.get(f, f))
            cb.setChecked(True)
            self.check_boxes[f] = cb
            layout.addWidget(cb)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.setLayout(layout)

    def selected_fields(self):
        return [f for f, cb in self.check_boxes.items() if cb.isChecked()]