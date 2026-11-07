# utils/file_utils.py
import os
import re
import sys
from utils.constants import ATTACHMENTS_DIR  # or from database.db_manager if defined there

def safe_filename_component(text):
    """
    Sanitizes free-text (e.g. a fax number) for safe use as part of a filename,
    stripping path separators so it can't be used to escape the target directory.
    """
    text = str(text).strip()
    text = text.replace(os.sep, '_')
    if os.altsep:
        text = text.replace(os.altsep, '_')
    return re.sub(r'[^\w.-]', '_', text) or '_'

def resource_path(path):
    meipass = getattr(sys, '_MEIPASS', None)
    if meipass:
        return os.path.join(meipass, path)
    return os.path.join(os.path.dirname(__file__), path)

def get_employee_folder(file_no):
    """
    Returns the path to the employee's attachment folder based on file_no.
    """
    file_no = str(file_no)
    if not file_no.isdigit():
        raise ValueError("file_no must contain digits only")
    folder = os.path.join(ATTACHMENTS_DIR, file_no)
    if not os.path.exists(folder):
        os.makedirs(folder)
    return folder
