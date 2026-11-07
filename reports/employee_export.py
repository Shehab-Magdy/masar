from PyQt5.QtWidgets import QMessageBox
from utils.constants import AR_LABELS
from reports.pdf_builder import MasarPDFReport, no_wrap_indices_for


class EmployeeReport:
    """Class responsible for generating employee PDF reports."""

    def __init__(self, conn):
        self.conn = conn

    def generate_filtered_pdf(self, selected_fields, rows, first_line_header, second_line_header,
                               bg_bytes, file_path, font_size=11, orientation="landscape", parent=None):
        headers = ["م"] + [AR_LABELS[f] for f in selected_fields]
        table_rows = []
        for idx, emp in enumerate(rows):
            row = [idx + 1]
            for i in range(len(selected_fields)):
                val = emp[i] if i < len(emp) and emp[i] else ""
                row.append(val)
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
            pdf.add_title("بيانات الموظفين المدنيين")
            pdf.add_zebra_table(headers, table_rows, no_wrap_indices_for(selected_fields))
            pdf.output(file_path)
            if parent:
                QMessageBox.information(parent, "تم", "تم تصدير النتائج بنجاح كملف PDF.")
            return True, "تم تصدير النتائج بنجاح كملف PDF."
        except Exception as e:
            if parent:
                QMessageBox.critical(parent, "خطأ", f"حدث خطأ أثناء التصدير: {e}")
            return False, str(e)
