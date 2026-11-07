# Masar Employee Management System

**مسار - نظام إدارة الموظفين**

Masar is a desktop application for managing employee records, attachments, and reports. It provides an easy-to-use Arabic interface for HR departments to store, search, and report on employee data.

---

## Features

- **Employee Management:** 
  - Add, edit, delete, and search employee records.
  - **New:** Manage relative/emergency contact information.
  - **Advanced Search:** Filter by name, department, file number, national ID, or specific combinations of fields.
- **Correspondence Management (المراسلات):** 
  - Track incoming and outgoing faxes and correspondence.
  - Attach images to correspondence records.
  - Search by date range, subject, or fax number.
  - Export correspondence search results to PDF.
- **Attachments:** Upload and manage files and personal photos for each employee.
- **Reports:** Generate basic and comprehensive employee reports.
- **PDF Export:** 
  - Export employee reports as RTL, Arabic-compatible PDF files.
  - **Customizable Columns:** Select exactly which columns to include in the PDF export.
- **Arabic Interface:** All labels and controls are in Arabic for local usability.
- **Modern GUI:** Built with PyQt5 for a modern user experience.
- **Arabic Text Normalization:** Search and save operations normalize Arabic text for better matching.
- **Custom App Icon:** Uses `masar.png` or `masar.ico` as the application icon.
- **Simple Deployment:** Runs as a standalone executable after packaging.

---

## Requirements

- **Python 3.7+**
- **PyQt5** (for the GUI)
- **fpdf2** + **uharfbuzz** (for PDF export with Arabic text shaping)
- **Pillow** (for image handling)
- **SQLite3** (included with Python)

Install all dependencies with:
```sh
pip install pyqt5 pillow fpdf2 uharfbuzz
```

---

## Usage

### Run Directly

To start the application, run:

```sh
python main.py
```

### Packaging as an Executable

You can package Masar as a standalone Windows executable using PyInstaller:

1. **Install PyInstaller**

   ```sh
   pip install pyinstaller
   ```

2. **Build the Executable**

   ```sh
   pyinstaller --onefile --windowed --name "Masar" --icon "assets/icons/masar.ico" --add-data "assets;assets" --add-data "config.json;." --add-data "utils/pdf_bg_utils.py;." --add-data "attachments;attachments" main.py
   ```

   The executable will be created in the `dist` directory.

3. **Copy Attachments Folder**

   Ensure the `attachments` folder is placed next to the executable for file storage.

4. **App Icon**

   The app uses `masar.png` or `masar.ico` as its icon. Make sure the icon file is present in the same directory as `main.py`.

---

## Application Structure

- **main.py**: Main application file.
- **masar.db**: SQLite database (created automatically).
- **attachments/**: Directory for storing uploaded files and photos.
- **ui/**: Contains UI tabs (Employee, Correspondence, Dashboard).
- **utils/**: Utility scripts and constants.

---

## How to Use

### Employee Management
1. **Add Employee:** Fills in employee details (including Relatives info) and click "إضافة".
2. **Edit Employee:** Select an employee from the list, modify details, and click "تعديل".
3. **Delete Employee:** Select an employee and click "حذف".
4. **Search:** 
   - **Quick Search:** Use the top search bar for Name, Dept, File No, or ID.
   - **Advanced Search:** Click "بحث بالمواصفات" to filter by multiple specific fields (like Degree, Address, etc.).
5. **Attachments:** Upload files and personal photos for each employee.
6. **PDF Export:** Click "تصدير النتائج كـ PDF" to select specific columns and generate a report.

### Correspondence (المراسلات)
1. **Navigate:** Go to the "المراسلات" tab.
2. **Add Entry:** Enter Fax Number, Date, Subject, Type (Import/Export), and other details. Click "إضافة".
3. **Attachments:** Upload images related to the fax.
4. **Search:** Filter by Subject, Fax Number, or Date Range.
5. **Export:** Click "تصدير النتائج" to save the current view as a PDF.

---

## Notes

- All data is stored locally in `masar.db`.
- Attachments are saved in the `attachments` directory.
- The application interface is in Arabic.
- Arabic text is normalized before saving and searching for better matching (e.g., "أ" and "ا" are treated the same).
- The PDF report is generated using fpdf2 and supports Arabic fonts and right-to-left layout via its built-in text-shaping engine. The Amiri font (`assets/Amiri-Regular.ttf`) is embedded directly into each PDF, so no system font installation is required.

---

## Troubleshooting

- If you encounter issues with images, ensure Pillow is installed.
- For database errors, delete `masar.db` to reset (all data will be lost).
- Always keep the `attachments` folder next to the executable.
- If PDF export fails, ensure `fpdf2` and `uharfbuzz` are installed and that `assets/Amiri-Regular.ttf` is present.

---

## License

This project is provided as-is for internal use.

---

## Contact

For support or suggestions, please contact the developer.

---

**تعليمات البناء بالعربية موجودة في نهاية ملف main.py**