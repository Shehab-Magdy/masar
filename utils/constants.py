
import os


DB_FILE = "masar.db"
ATTACHMENTS_DIR = "attachments"

AR_LABELS = {
    "name": "الاسم",
    "grade": "الدرجة",
    "grade_date": "تاريخ الحصول عليها",
    "hire_date": "تاريخ التعيين",
    "file_no": "رقم الملف",
    "qualification": "المؤهل",
    "functional_group": "مجموعة وظيفية",
    "type_group": "مجموعة نوعية",
    "job_title": "المسمى الوظيفي",
    "department": "القسم",
    "current_work": "العمل القائم به",
    "birth_date": "تاريخ الميلاد",
    "insurance_no": "رقم تأميني",
    "national_id": "رقم قومي",
    "address": "عنوان حالي",
    "phone": "رقم التليفون",
    "notes": "ملاحظات",
    "attachments": "ملفات مرتبطة",
    "personal_photo": "صورة شخصية",
    "retirement_date": "تاريخ المعاش"
    ,"insurance_doc": "وثيقة التامين"
    ,"serial": "مسلسل"
}

EMPLOYEE_FIELDS = [
    "name", "grade", "grade_date", "hire_date", "file_no", "qualification",
    "functional_group", "type_group", "job_title", "department", "current_work",
    "birth_date", "retirement_date", "insurance_no", "national_id", "address", "phone", "insurance_doc", "notes"
]

bg_path = os.path.join(os.getcwd(), 'masar-bg.png')
cfg_path = os.path.join(os.getcwd(), 'config.json')