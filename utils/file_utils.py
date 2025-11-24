# utils/file_utils.py
import os
import sys
from utils.constants import ATTACHMENTS_DIR  # or from database.db_manager if defined there

def resource_path(path):
    meipass = getattr(sys, '_MEIPASS', None)
    if meipass:
        return os.path.join(meipass, path)
    return os.path.join(os.path.dirname(__file__), path)

def get_employee_folder(file_no):
    """
    Returns the path to the employee's attachment folder based on file_no.
    """
    folder = os.path.join(ATTACHMENTS_DIR, str(file_no))
    if not os.path.exists(folder):
        os.makedirs(folder)
    return folder
