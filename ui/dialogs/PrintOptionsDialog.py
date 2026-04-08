from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QCheckBox,
    QDialogButtonBox, QScrollArea, QWidget
)
from PyQt5.QtCore import Qt


class PrintOptionsDialog(QDialog):
    """Dialog for selecting print options before generating PDF reports."""
    
    def __init__(self, parent=None, default_landscape=True):
        """
        Initialize the Print Options Dialog.
        
        :param parent: Parent widget
        :param default_landscape: Whether the dialog defaults to landscape mode
        """
        super().__init__(parent)
        self.setWindowTitle("خيارات الطباعة")
        self.setGeometry(100, 100, 400, 200)
        self.landscape_mode = default_landscape
        self.default_landscape = default_landscape
        self.init_ui()
        
    def init_ui(self):
        """Initialize the UI components."""
        layout = QVBoxLayout()
        
        # Title
        title_label = QLabel("خيارات الطباعة")
        title_label.setStyleSheet("font-weight: bold; font-size: 13px;")
        layout.addWidget(title_label)
        
        # Adding a scroll area for future expandability
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout()
        scroll.setWidget(scroll_widget)
        
        # Landscape mode checkbox (default checked or not based on report type)
        landscape_layout = QHBoxLayout()
        self.landscape_checkbox = QCheckBox("طباعة أفقية (Landscape)")
        self.landscape_checkbox.setChecked(self.default_landscape)
        self.landscape_checkbox.stateChanged.connect(self.on_landscape_changed)
        landscape_layout.addWidget(self.landscape_checkbox)
        landscape_layout.addStretch()
        scroll_layout.addLayout(landscape_layout)
        
        scroll_widget.setLayout(scroll_layout)
        layout.addWidget(scroll)
        
        # Dialog buttons
        button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
        self.setLayout(layout)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        
    def on_landscape_changed(self, state):
        """Handle landscape checkbox state change."""
        self.landscape_mode = self.landscape_checkbox.isChecked()
    
    def is_landscape(self):
        """
        Get the selected orientation.
        
        :return: True if landscape, False if portrait
        :rtype: bool
        """
        return self.landscape_checkbox.isChecked()
