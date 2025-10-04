# Masar Employee Management System

**مسار - نظام إدارة الموظفين**

Masar is a desktop application for managing employee records, attachments, and reports. It provides an easy-to-use Arabic interface for HR departments to store, search, and report on employee data.

---

## Features

- **Employee Management:** Add, edit, delete, and search employee records.
- **Attachments:** Upload and manage files and personal photos for each employee.
- **Reports:** Generate basic and comprehensive employee reports.
- **Arabic Interface:** All labels and controls are in Arabic for local usability.
- **Modern GUI:** Built with PyQt5 for a modern user experience.
- **PDF Export:** Export employee reports as RTL, Arabic-compatible PDF files using WeasyPrint.
- **Arabic Text Normalization:** Search and save operations normalize Arabic text for better matching.
- **Custom App Icon:** Uses `masar.png` or `masar.ico` as the application icon.
- **Simple Deployment:** Runs as a standalone executable after packaging.

---

## Requirements

- **Python 3.7+**
- **PyQt5** (for the GUI)
- **WeasyPrint** (for PDF export)
- **Pillow** (for image handling)
- **SQLite3** (included with Python)

Install all dependencies with:
```sh
pip install pyqt5 pillow weasyprint
```

---

## Installation

1. **Clone or Download the Repository**

   Download the source code and place it in your desired directory.

2. **Install Dependencies**

   Open a terminal and run:

   ```sh
   pip install pyqt5 pillow weasyprint
   ```

   > **Note:** On some systems, WeasyPrint may require additional system libraries. See [WeasyPrint installation docs](https://weasyprint.readthedocs.io/en/stable/install.html).

---

## Usage

### Run Directly

To start the application, run:

```sh
python masar.py
```

### Packaging as an Executable

You can package Masar as a standalone Windows executable using PyInstaller:

1. **Install PyInstaller**

   ```sh
   pip install pyinstaller
   ```

2. **Build the Executable**

   ```sh
   pyinstaller --onefile --windowed --icon "masar.ico" --add-data "masar-bg.png;." --add-data "Amiri-Regular.ttf;." --add-data "config.json;." --add-data "pdf_bg_utils.py;." --add-data "attachments;attachments" masar.py   
   ```

   The executable will be created in the `dist` directory.

3. **Copy Attachments Folder**

   Ensure the `attachments` folder is placed next to the executable for file storage.

4. **App Icon**

   The app uses `masar.png` or `masar.ico` as its icon. Make sure the icon file is present in the same directory as `masar.py`.

---

## Application Structure

- **masar.py**: Main application file.
- **masar.db**: SQLite database (created automatically).
- **attachments/**: Directory for storing uploaded files and photos.
- **masar.png / masar.ico**: Application icon.

---

## How to Use

1. **Add Employee:** Fill in the employee details and click "إضافة".
2. **Edit Employee:** Select an employee from the list, modify details, and click "تعديل".
3. **Delete Employee:** Select an employee and click "حذف".
4. **Search:** Use the search field at the top to filter employees by name, department, file number, or national ID. Search is normalized for Arabic text.
5. **Attachments:** Upload files and personal photos for each employee.
6. **Reports:** Go to the "التقارير" tab to generate and view reports.
7. **Export PDF:** Click "تصدير كـ PDF" to export a landscape, RTL, Arabic-compatible PDF report. The file name will be generated as `Employees yyyy-mm-dd hh-mm-ss.pdf`.

---

## Notes

- All data is stored locally in `masar.db`.
- Attachments are saved in the `attachments` directory.
- The application interface is in Arabic.
- Arabic text is normalized before saving and searching for better matching (e.g., "أ" and "ا" are treated the same).
- The PDF report is generated using WeasyPrint and supports Arabic fonts and right-to-left layout. Make sure you have a suitable Arabic font (like Amiri) available, or adjust the CSS in the code as needed.

---

## Troubleshooting

- If you encounter issues with images, ensure Pillow is installed.
- For database errors, delete `masar.db` to reset (all data will be lost).
- Always keep the `attachments` folder next to the executable.
- If PDF export fails, ensure WeasyPrint and its dependencies are installed and that you have a suitable Arabic font.

---

## License

This project is provided as-is for internal use.

---

## Contact

For support or suggestions, please contact the developer.

---

**تعليمات البناء بالعربية موجودة في نهاية ملف masar.py**