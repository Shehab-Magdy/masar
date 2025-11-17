"""
pytest configuration and PyQt5 mocking for all tests.
This file is auto-discovered by pytest and runs before any tests.
"""

import sys
from unittest.mock import MagicMock

# Mock ALL PyQt5 and external dependencies BEFORE anything else imports them
sys.modules['PyQt5'] = MagicMock()
sys.modules['PyQt5.QtWidgets'] = MagicMock()
sys.modules['PyQt5.QtGui'] = MagicMock()
sys.modules['PyQt5.QtCore'] = MagicMock()
sys.modules['weasyprint'] = MagicMock()
sys.modules['pdf_bg_utils'] = MagicMock()
