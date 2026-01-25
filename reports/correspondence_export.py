import os
import datetime
from PyQt5.QtWidgets import (QFileDialog, QMessageBox
)
from weasyprint import HTML, CSS
import base64
from utils.constants import bg_path, cfg_path
from utils.pdf_bg_utils import process_bg_image

def export_correspondence_pdf(parent, ids):
    if not ids:
        QMessageBox.warning(parent, "تنبيه", "لا توجد نتائج للتصدير")
        return
    c = parent.conn.cursor()
    # Prepare placeholders for SQL IN clause
    placeholders = ','.join('?' for _ in ids)
    q = f"""SELECT id, fax_number, correspondence_type, fax_date, from_person, to_person, subject, notes, image_path
            FROM correspondence WHERE id IN ({placeholders}) ORDER BY fax_date DESC"""
    c.execute(q, ids)
    rows = c.fetchall()
    if not rows:
        QMessageBox.warning(parent, "تنبيه", "لا توجد نتائج للتصدير")
        return

    now = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    file_path, _ = QFileDialog.getSaveFileName(parent, "حفظ المراسلات كـ PDF", f"correspondence_{now}.pdf", "PDF Files (*.pdf)")
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

    # build html
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
            {'background: url("'+bg_url+'") no-repeat center center; background-size: contain;' if bg_url else ''} }}
        table {{ 
            border-collapse: collapse; 
            width: 100%; 
            table-layout: auto; 
        }}
        th, td {{ 
            border: 1px solid #888; 
            padding: 6px; 
            vertical-align: top; 
            text-align: right; 
        }}
        th:nth-child(-n+6), td:nth-child(-n+6) {{
            width: 1%;
            white-space: nowrap;
        }}
        th:nth-child(n+7), td:nth-child(n+7) {{
            white-space: pre-wrap;
            word-break: break-word;
        }}
        th {{ 
            background: #b3d1f7; 
        }}
        tr {{
            page-break-inside: avoid;
        }}
        tr:nth-child(odd) {{ 
            background-color: transparent; 
        }}
        tr:nth-child(even) {{ 
            background-color: #f2f2f2; 
        }}
        @page {{ 
            size: A4 portrait; 
            margin: 1cm 0.5cm 1.5cm 0.5cm;
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
        <h2 style="font-size:15px; text-align:center;">قائمة المراسلات</h2>
        <table dir="rtl">
        <thead>
            <tr>
            <th>م</th>
            <th>رقم الفاكس</th>
            <th>نوع المراسلة</th>
            <th>التاريخ</th>
            <th>من</th>
            <th>إلى</th>
            <th>الموضوع</th>
            <th>الملاحظات</th>
            </tr>
        </thead>
        <tbody>
    """

    for idx, r in enumerate(rows):
        serial = idx + 1
        # r[1]=fax_no, r[2]=corr_type, r[3]=fax_date, r[4]=from_p, r[5]=to_p, r[6]=subj, r[7]=notes
        html += f"<tr><td>{serial}</td><td>{r[1] or ''}</td><td>{r[2] or ''}</td><td>{r[3] or ''}</td><td>{r[4] or ''}</td><td>{r[5] or ''}</td><td>{(r[6] or '')}</td><td>{(r[7] or '').replace('\n','<br/>')}</td></tr>"

    html += """
        </tbody>
        </table>
    </body>
    </html>
    """

    try:
        css = CSS(string='@page { size: A4 portrait; margin: 1cm; }')
        HTML(string=html, base_url=os.getcwd()).write_pdf(file_path, stylesheets=[css])
        QMessageBox.information(parent, "تم", "تم تصدير النتائج بنجاح كملف PDF.")
    except Exception as e:
        QMessageBox.critical(parent, "خطأ", f"حدث خطأ أثناء التصدير: {e}")
