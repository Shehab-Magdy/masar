# utils/config_loader.py
import json
import os
from utils.constants import cfg_path

DEFAULT_FONT_SIZE = 11


def load_report_config():
    """
    Loads report header/style settings from config.json.

    Returns (first_line_header, second_line_header, font_size), falling back to
    safe defaults if the file is missing or invalid.
    """
    first_line_header = ""
    second_line_header = ""
    font_size = DEFAULT_FONT_SIZE
    if os.path.isfile(cfg_path):
        try:
            with open(cfg_path, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
            first_line_header = cfg.get('firstLineHeader', "")
            second_line_header = cfg.get('secondLineHeader', "")
            font_size = cfg.get('font-size', DEFAULT_FONT_SIZE)
        except Exception:
            first_line_header = ""
            second_line_header = ""
            font_size = DEFAULT_FONT_SIZE
    return first_line_header, second_line_header, font_size
