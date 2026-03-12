"""
pytest configuration and PyQt5 mocking for all tests.
This file is auto-discovered by pytest and runs before any tests.
"""

import sys
from unittest.mock import MagicMock

# Mock PyQt5 and UI-related dependencies BEFORE anything else imports them
sys.modules['PyQt5'] = MagicMock()
sys.modules['PyQt5.QtWidgets'] = MagicMock()
sys.modules['PyQt5.QtGui'] = MagicMock()
sys.modules['PyQt5.QtCore'] = MagicMock()
# Don't mock weasyprint - let tests patch it as needed
sys.modules['pdf_bg_utils'] = MagicMock()
