# reports/retire_export.py
import datetime
import calendar
import os

from PyQt5.QtWidgets import QMessageBox, QFileDialog

from utils.constants import AR_LABELS, EMPLOYEE_FIELDS, cfg_path, bg_path
from utils.config_loader import load_report_config
from utils.pdf_bg_utils import process_bg_image
from reports.pdf_builder import MasarPDFReport, no_wrap_indices_for


def export_retire_pdf(conn, months=6, parent=None, orientation="landscape"):
    """
    Export the full data of employees whose retirement date is within the next 'months' months as a PDF,
    using the same columns and design as the main employee export.
    """
    # Calculate date range
    today = datetime.date.today()
    # The period includes the current month as its first month (handle year wrap)
    start_date = today.replace(day=1)
    year = today.year + (today.month + months - 2) // 12
    month = (today.month + months - 2) % 12 + 1
    last_day = calendar.monthrange(year, month)[1]
    end_date = datetime.date(year, month, last_day)

    report_fields = EMPLOYEE_FIELDS[:-3]
    headers = ["م"] + [AR_LABELS[f] for f in report_fields]

    # Query all fields for employees retiring within the next N months
    c = conn.cursor()
    c.execute(f"""
        SELECT {', '.join(EMPLOYEE_FIELDS)} FROM employee
        WHERE retirement_date IS NOT NULL AND retirement_date != ''
        ORDER BY retirement_date
    """)
    all_rows = c.fetchall()

    # Filter in Python
    rows = []
    for emp in all_rows:
        date_str = emp[EMPLOYEE_FIELDS.index("retirement_date")]
        try:
            # Try both formats
            try:
                dt = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
            except ValueError:
                dt = datetime.datetime.strptime(date_str, "%d/%m/%Y").date()
            if start_date <= dt <= end_date:
                rows.append(emp)
        except Exception:
            continue

    if not rows:
        QMessageBox.warning(parent, "تنبيه", "لا يوجد بيانات للتصدير.")
        return

    now = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    default_name = f"Employees_Retirement_{now}.pdf"
    file_path, _ = QFileDialog.getSaveFileName(
        parent,
        "حفظ قائمة المعاش كـ PDF",
        default_name,
        "PDF Files (*.pdf)"
        )
    if not file_path:
        return

    first_line_header, second_line_header, font_size = load_report_config()

    bg_bytes = None
    if os.path.isfile(bg_path) and os.path.isfile(cfg_path):
        try:
            bg_bytes = process_bg_image(bg_path, cfg_path)
        except Exception:
            bg_bytes = None

    table_rows = []
    for idx, emp in enumerate(rows):
        row = [idx + 1]
        for i in range(len(report_fields)):
            row.append(emp[i] if emp[i] else "")
        table_rows.append(row)

    try:
        pdf = MasarPDFReport(
            orientation=orientation,
            font_size=font_size,
            first_line_header=first_line_header,
            second_line_header=second_line_header,
            bg_bytes=bg_bytes,
        )
        pdf.add_page()
        pdf.add_title(f"بيانات الموظفين الذين تاريخ معاشهم خلال {months} أشهر القادمة")
        pdf.add_zebra_table(headers, table_rows, no_wrap_indices_for(report_fields))
        pdf.output(file_path)
        QMessageBox.information(parent, "تم", "تم تصدير القائمة بنجاح كملف PDF.")
    except Exception as e:
        QMessageBox.critical(parent, "خطأ", f"حدث خطأ أثناء تصدير القائمة: {e}")
