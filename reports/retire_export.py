# reports/retire_export.py
import datetime, calendar, os
from weasyprint import HTML, CSS
from utils.file_utils import resource_path
from PyQt5.QtWidgets import QMessageBox, QFileDialog
from database.db_manager import EMPLOYEE_FIELDS
from utils.constants import AR_LABELS, EMPLOYEE_FIELDS, cfg_path, bg_path
import base64

from utils.pdf_bg_utils import process_bg_image
def export_retire_pdf(conn, months=6, parent=None):
    """
    Export the full data of employees whose retirement date is within the next 'months' months as a PDF,
    using the same columns and design as the main employee export.
    """
    # Calculate date range
    today = datetime.date.today()
    # Add months to today (handle year wrap)
    year = today.year + (today.month + months - 1) // 12
    month = (today.month + months - 1) % 12 + 1
    last_day = calendar.monthrange(year, month)[1]
    end_date = datetime.date(year, month, last_day)
    
    headers = ["م"] + [AR_LABELS[f] for f in EMPLOYEE_FIELDS]

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
            if today <= dt <= end_date:
                rows.append(emp)
        except Exception:
            continue
    
    if not rows:
        QMessageBox.warning(parent, "تنبيه", "لا يوجد بيانات لتصديرها.")
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
    
    # Prepare background image as base64 (only if file and config exist and are valid)
    bg_url = None
    first_line_header = ""
    second_line_header = ""
    if os.path.isfile(cfg_path):
        try:
            import json
            with open(cfg_path, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
            first_line_header = cfg.get('firstLineHeader', "")
            second_line_header = cfg.get('secondLineHeader', "")
            font_size = cfg.get('font-size', 11)
        except Exception:
            first_line_header = ""
            second_line_header = ""
    if os.path.isfile(bg_path) and os.path.isfile(cfg_path):
        try:
            bg_bytes = process_bg_image(bg_path, cfg_path)
            bg_b64 = base64.b64encode(bg_bytes).decode('utf-8')
            bg_url = f"data:image/png;base64,{bg_b64}"
        except Exception:
            bg_url = None
    
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
                margin-bottom: 20px;
            }}
            th, td {{
                border: 1px solid #888;
                padding: 6px 4px;
                word-break: break-word;
                vertical-align: top;
                text-align: right;
            }}
            th {{
                background: #b3d1f7;
            }}
            tr:nth-child(odd) {{
                background-color: transparent;
            }}
            tr:nth-child(even) {{
                background-color: #f2f2f2;
            }}
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
        <h2 style="font-size:15px; text-align:center;">بيانات الموظفين الذين تاريخ معاشهم خلال {months} أشهر القادمة</h2>
        <table dir="rtl">
            <thead>
                <tr>
                    {''.join(f'<th>{h}</th>' for h in headers)}
                </tr>
            </thead>
            <tbody>
    """

    for idx, emp in enumerate(rows):
        serial = idx + 1
        html += f'<tr><td>{serial}</td>'
        for i, f in enumerate(EMPLOYEE_FIELDS):
            val = emp[i] if emp[i] else ""
            html += f'<td>{val}</td>'
        html += '</tr>'

    html += """
        </tbody>
    </table>
</body>
</html>
"""

    try:
        css = CSS(string='''
            @page { size: A4 landscape; margin: 1cm 0.5cm 1.5cm 0.5cm; }
        ''')
        HTML(string=html, base_url=os.getcwd()).write_pdf(file_path, stylesheets=[css])
        QMessageBox.information(parent, "تم", "تم تصدير القائمة بنجاح كملف PDF.")
    except Exception as e:
        QMessageBox.critical(parent, "خطأ", f"حدث خطأ أثناء تصدير القائمة: {e}")
