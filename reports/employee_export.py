import os
import datetime
import base64
from weasyprint import HTML, CSS
from PyQt5.QtWidgets import QMessageBox
from utils.constants import AR_LABELS, EMPLOYEE_FIELDS
from utils.pdf_bg_utils import process_bg_image


class EmployeeReport:
    """Class responsible for generating employee PDF reports."""
    no_wrap_fields = {"file_no","grade_date","hire_date","birth_date","retirement_date","insurance_no","national_id","phone","relative_phone"}

    def __init__(self, conn):
        self.conn = conn

    def generate_filtered_pdf(self, selected_fields, rows, first_line_header, second_line_header, bg_url, file_path, font_size=11, parent=None):
        # build headers html with classes where needed
        headers_html = ''
        headers_html += '<th class="no-wrap">م</th>'
        for f in selected_fields:
            cls = ' class="no-wrap"' if f in self.no_wrap_fields else ''
            headers_html += f'<th{cls}>' + AR_LABELS[f] + '</th>'
        # build colgroup so no-wrap cols can shrink to content
        colgroup_html = '<colgroup>'
        colgroup_html += '<col class="no-wrap" />'  # serial
        for f in selected_fields:
            colgroup_html += '<col class="no-wrap" />' if f in self.no_wrap_fields else '<col />'
        colgroup_html += '</colgroup>'

        html = f"""
        <html lang="ar">
        <head>
            <meta charset="utf-8">
            <style>
                @font-face {{
                    font-family: 'Amiri';
                    src: url('assets/Amiri-Regular.ttf') format('truetype');
                }}
                body {{
                    direction: rtl;
                    font-family: 'Amiri', 'Cairo', 'Tahoma', sans-serif;
                    font-size: {font_size}px;
                    {'background: url("'+bg_url+'"); background-size: contain; background-repeat: no-repeat; background-position: center center;' if bg_url else ''}
                }}
                table {{
                    border-collapse: collapse;
                    width: 100%;
                    max-width: 100%;
                    table-layout: auto;
                    box-sizing: border-box;
                    margin-bottom: 20px;
                }}
                th, td {{
                    border: 1px solid #888;
                    padding: 6px 4px;
                    white-space: pre-wrap;
                    overflow-wrap: break-word;
                    vertical-align: top;
                    text-align: right;
                    box-sizing: border-box;
                }}
                .no-wrap {{ white-space: nowrap; }}
                td.no-wrap {{ padding-left: 6px; padding-right: 6px; }}
                col.no-wrap {{ width: auto; }}
                th {{
                    background: #b3d1f7;
                }}
                tr {{
                    page-break-inside: avoid;
                }}
                /* light rows transparent so the PDF background shows through */
                tr.zebra1 {{ background-color: transparent; }}
                tr.zebra2 {{ background-color: #f2f2f2; }}
                @page {{
                    size: A4 landscape;
                    margin: 1cm 1cm 2cm 1cm;
                    @bottom-center {{
                        content: "الصفحة " counter(page) " من " counter(pages);
                        font-family: 'Amiri', 'Cairo', 'Tahoma', sans-serif;
                        font-size: 12px;
                        color: #444;
                    }}
                }}
            </style>
        </head>
        <body>
            <div style="text-align:right; margin-bottom: 8px;">
                <div style="font-size:13px; color:#1976d2;">{first_line_header}</div>
                <div style="font-size:13px; color:#1976d2;">{second_line_header}</div>
            </div>
            <h2 style="font-size:15px; text-align:center;">بيانات الموظفين المدنيين في الورش الرئيسية للطائرات</h2>
            <table dir="rtl">
                {colgroup_html}
                <thead>
                    <tr>
                        {headers_html}
                    </tr>
                </thead>
                <tbody>
        """

        for idx, emp in enumerate(rows):
            row_class = "zebra1" if idx % 2 == 0 else "zebra2"
            serial = idx + 1
            # build cells with no-wrap for specific fields
            cells = [f'<td class="no-wrap">{serial}</td>']
            for i, f_field in enumerate(selected_fields):
                val = emp[i] if i < len(emp) and emp[i] else ""
                cls = ' class="no-wrap"' if f_field in self.no_wrap_fields else ''
                cells.append(f'<td{cls}>' + str(val) + '</td>')
            html += f'<tr class="{row_class}">' + ''.join(cells) + '</tr>'

        html += """
                </tbody>
            </table>
        </body>
        </html>
        """

        try:
            css = CSS(string='@page { size: A4 landscape; margin: 1cm 0.5cm 1.5cm 0.5cm; }')
            HTML(string=html, base_url=os.getcwd()).write_pdf(file_path, stylesheets=[css])
            if parent:
                QMessageBox.information(parent, "تم", "تم تصدير النتائج بنجاح كملف PDF.")
            return True, "تم تصدير النتائج بنجاح كملف PDF."
        except Exception as e:
            if parent:
                QMessageBox.critical(parent, "خطأ", f"حدث خطأ أثناء التصدير: {e}")
            return False, str(e)

    def generate_full_pdf(self, employees, first_line_header, second_line_header, bg_url, file_path, font_size=9, parent=None):
        # build headers html with classes
        headers_html = ''
        headers_html += '<th class="no-wrap">م</th>'
        for f in EMPLOYEE_FIELDS:
            cls = ' class="no-wrap"' if f in self.no_wrap_fields else ''
            headers_html += f'<th{cls}>' + AR_LABELS[f] + '</th>'
        # build colgroup so no-wrap cols can shrink to content
        colgroup_html = '<colgroup>'
        colgroup_html += '<col class="no-wrap" />'  # serial
        for f in EMPLOYEE_FIELDS:
            colgroup_html += '<col class="no-wrap" />' if f in self.no_wrap_fields else '<col />'
        colgroup_html += '</colgroup>'

        html = f"""
        <html lang="ar">
        <head>
            <meta charset="utf-8">
            <style>
                @font-face {{
                    font-family: 'Amiri';
                    src: url('Amiri-Regular.ttf');
                }}
                body {{
                    direction: rtl;
                    font-family: 'Amiri', 'Cairo', 'Tahoma', sans-serif;
                    font-size: {font_size}px;
                    {'background: url("'+bg_url+'"); background-size: contain; background-repeat: no-repeat; background-position: center center;' if bg_url else ''}
                }}
                table {{
                    border-collapse: collapse;
                    width: 100%;
                    max-width: 100%;
                    table-layout: auto;
                    box-sizing: border-box;
                    margin-bottom: 20px;
                }}
                th, td {{
                    border: 1px solid #888;
                    padding: 6px 4px;
                    white-space: pre-wrap;
                    overflow-wrap: break-word;
                    vertical-align: top;
                    text-align: right;
                    box-sizing: border-box;
                }}
                .no-wrap {{ white-space: nowrap; }}
                td.no-wrap {{ padding-left: 6px; padding-right: 6px; }}
                col.no-wrap {{ width: auto; }}
                th {{
                    background: #b3d1f7;
                }}
                tr {{
                    page-break-inside: avoid;
                }}
                /* light rows should be transparent so the PDF background shows through */
                tr:nth-child(odd) {{
                    background-color: transparent;
                }}
                tr:nth-child(even) {{
                    background-color: #f2f2f2;
                }}
                @page {{
                    size: A4 landscape;
                    margin: 1cm 0.5cm 1.5cm 0.5cm;
                    @top-right {{
                        content: '""" + first_line_header + """\\A""" + second_line_header + """';
                        font-size: 15px;
                        color: #1976d2;
                        text-align: right;
                        white-space: pre;
                  }}
                }}
            </style>
        </head>
        <body>
            <h2 style="text-align:center;">بيانات الموظفين المدنيين في الورش الرئيسية للطائرات</h2>
            <table dir="rtl">
                {colgroup_html}
                <thead>
                    <tr>
                        {headers_html}
                    </tr>
                </thead>
                <tbody>
        """

        for idx, emp in enumerate(employees):  # or employees
            row_class = "zebra1" if idx % 2 == 0 else "zebra2"
            serial = idx + 1
            cells = [f'<td class="no-wrap">{serial}</td>']
            for i, f_field in enumerate(EMPLOYEE_FIELDS):
                val = emp[i] if i < len(emp) and emp[i] else ""
                cls = ' class="no-wrap"' if f_field in self.no_wrap_fields else ''
                cells.append(f'<td{cls}>' + str(val) + '</td>')
            html += f'<tr class="{row_class}">' + ''.join(cells) + '</tr>'

        html += """
                </tbody>
            </table>
        </body>
        </html>
        """

        try:
            css = CSS(string='@page { size: A4 landscape; margin: 1cm 0.5cm 1.5cm 0.5cm; }')
            HTML(string=html, base_url=os.getcwd()).write_pdf(file_path, stylesheets=[css])
            if parent:
                QMessageBox.information(parent, "تم", "تم تصدير التقرير بنجاح كملف PDF.")
            return True, "تم تصدير التقرير بنجاح كملف PDF."
        except Exception as e:
            if parent:
                QMessageBox.critical(parent, "خطأ", f"حدث خطأ أثناء التصدير: {e}")
            return False, str(e)
