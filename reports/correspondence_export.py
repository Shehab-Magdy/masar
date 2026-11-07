# reports/correspondence_export.py
from PyQt5.QtWidgets import QMessageBox
from reports.pdf_builder import MasarPDFReport

HEADERS = ["مسلسل", "رقم الفاكس", "التاريخ", "من", "إلى", "الموضوع", "الملاحظات"]
NO_WRAP_COLS = {0, 1, 2}  # مسلسل، رقم الفاكس، التاريخ: fixed short values


def export_correspondence_pdf(rows, first_line_header, second_line_header, bg_bytes,
                               file_path, font_size=11, orientation="portrait", parent=None):
    """
    rows: list of (fax_number, fax_date, from_person, to_person, subject, notes, image_path)
    tuples, as fetched from the correspondence table.
    """
    table_rows = []
    for idx, r in enumerate(rows):
        fax_no, fax_date, from_p, to_p, subj, notes, _img = r
        table_rows.append([idx + 1, fax_no or "", fax_date or "", from_p or "", to_p or "", subj or "", notes or ""])

    try:
        pdf = MasarPDFReport(
            orientation=orientation,
            font_size=font_size,
            first_line_header=first_line_header,
            second_line_header=second_line_header,
            bg_bytes=bg_bytes,
        )
        pdf.add_page()
        pdf.add_title("قائمة المراسلات")
        pdf.add_zebra_table(HEADERS, table_rows, NO_WRAP_COLS)
        pdf.output(file_path)
        if parent:
            QMessageBox.information(parent, "تم", "تم تصدير النتائج بنجاح كملف PDF.")
        return True, "تم تصدير النتائج بنجاح كملف PDF."
    except Exception as e:
        if parent:
            QMessageBox.critical(parent, "خطأ", f"حدث خطأ أثناء التصدير: {e}")
        return False, str(e)
