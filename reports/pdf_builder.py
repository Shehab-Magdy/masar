# reports/pdf_builder.py
"""
Shared PDF-building base for the employee/retirement/correspondence reports.
Replaces the old WeasyPrint HTML/CSS templates with fpdf2, which is pure
Python (no native Pango/cairo dependency) and has built-in Arabic
contextual-shaping + RTL support via set_text_shaping().
"""
import io
import os

from fpdf import FPDF
from fpdf.fonts import FontFace

# cwd-relative, matching utils/constants.py's bg_path/cfg_path convention
# (resource_path() resolves relative to utils/, not the project root, so it
# doesn't apply here).
AMIRI_FONT_PATH = os.path.join(os.getcwd(), "assets", "Amiri-Regular.ttf")

HEADER_FILL = (179, 209, 247)  # #b3d1f7, matches the old report header shading
ZEBRA_FILL = (242, 242, 242)  # #f2f2f2, matches the old alternating-row shading

# Fields that held their own value short (dates, numbers, ids) in the old
# WeasyPrint "no-wrap" CSS class; these are sized to fit their content on one
# line in the fpdf2 tables below, instead of wrapping like free-text columns.
NO_WRAP_FIELDS = {
    "file_no", "grade_date", "hire_date", "birth_date", "retirement_date",
    "insurance_no", "national_id", "phone", "relative_phone",
}

CELL_PADDING_MM = 4  # extra room beyond measured text width, for a no-wrap column
# Safety floor for a wrapping column, in mm - just enough to fit a single
# Arabic character so fpdf2 can still lay out wrapped lines; not an aesthetic
# target. Reports with many columns (e.g. the full retirement export) can
# still end up with narrow, heavily-wrapped text columns - that's expected
# once the no-wrap columns (dates/numbers) have claimed the room they need.
MIN_FLEX_WIDTH_MM = 6


def no_wrap_indices_for(fields, serial=True):
    """0-based column indices (within a header row that starts with a serial
    column, if any) that should stay on a single line rather than wrap."""
    offset = 1 if serial else 0
    indices = set(range(offset))
    indices.update(i + offset for i, f in enumerate(fields) if f in NO_WRAP_FIELDS)
    return indices


class MasarPDFReport(FPDF):
    """
    A portrait/landscape A4 report with:
      - Amiri font + Arabic text shaping enabled
      - an optional full-page watermark background image, repeated per page
      - a two-line report header (top), and a page-number footer (bottom)
      - "موجه الى" shown once, bottom-right of the first page only
    """

    def __init__(self, orientation="portrait", font_size=11,
                 first_line_header="", second_line_header="", bg_bytes=None):
        super().__init__(orientation=orientation, unit="mm", format="A4")
        self.first_line_header = first_line_header
        self.second_line_header = second_line_header
        self.bg_bytes = bg_bytes
        self.font_size = font_size
        self.set_margins(10, 20, 10)
        self.set_auto_page_break(auto=True, margin=18)
        self.alias_nb_pages()
        self.add_font("Amiri", "", AMIRI_FONT_PATH)
        self.set_text_shaping(use_shaping_engine=True)
        self.set_font("Amiri", size=font_size)

    def header(self):
        if self.bg_bytes:
            self.image(io.BytesIO(self.bg_bytes), x=0, y=0, w=self.w, h=self.h)
        self.set_font("Amiri", size=13)
        self.set_xy(self.l_margin, 8)
        if self.first_line_header:
            self.cell(w=self.epw, text=self.first_line_header, align="R")
        self.set_xy(self.l_margin, 14)
        if self.second_line_header:
            self.cell(w=self.epw, text=self.second_line_header, align="R")
        self.set_font("Amiri", size=self.font_size)
        self.set_xy(self.l_margin, self.t_margin)

    def footer(self):
        self.set_font("Amiri", size=10)
        self.set_y(-15)
        self.cell(w=self.epw, text=f"الصفحة {self.page_no()} من {{nb}}", align="C")
        if self.page_no() == 1:
            self.set_xy(self.l_margin, -10)
            self.set_font("Amiri", size=self.font_size)
            self.cell(w=self.epw, text="موجه الى", align="R")

    def add_title(self, text):
        self.set_font("Amiri", size=15)
        self.cell(w=self.epw, text=text, align="C", new_x="LMARGIN", new_y="NEXT")
        self.ln(2)
        self.set_font("Amiri", size=self.font_size)

    def _column_widths(self, headers, rows, no_wrap_cols):
        """
        Absolute mm width per column (pre-reversal order), summing to the
        full page width (self.epw) as fpdf2's table() normalizes col_widths
        proportionally to that width anyway: no_wrap columns get just enough
        to fit their widest actual value on one line, and the rest is split
        evenly across the remaining (wrapping) columns.
        """
        n = len(headers)
        fixed = {}
        for i in range(n):
            if i not in no_wrap_cols:
                continue
            width = self.get_string_width(str(headers[i]))
            for row in rows:
                val = row[i]
                text = str(val) if val is not None else ""
                width = max(width, self.get_string_width(text))
            fixed[i] = width + CELL_PADDING_MM

        flex_cols = [i for i in range(n) if i not in fixed]
        available = self.epw
        fixed_total = sum(fixed.values())

        if flex_cols:
            min_flex_total = MIN_FLEX_WIDTH_MM * len(flex_cols)
            if fixed_total > 0 and available - fixed_total < min_flex_total:
                shrink_to = max(available - min_flex_total, 0)
                scale = shrink_to / fixed_total if fixed_total else 0
                # never shrink a no-wrap column below the floor either - a
                # too-small scale here would otherwise make dates/numbers
                # wrap (defeating the point) or even crash fpdf2's renderer.
                scale = max(scale, MIN_FLEX_WIDTH_MM / max(fixed.values()))
                fixed = {i: w * scale for i, w in fixed.items()}
                fixed_total = sum(fixed.values())
            flex_width = max((available - fixed_total) / len(flex_cols), MIN_FLEX_WIDTH_MM * 0.1)
        elif fixed_total > available:
            scale = available / fixed_total
            fixed = {i: w * scale for i, w in fixed.items()}

        return [fixed[i] if i in fixed else flex_width for i in range(n)]

    def add_zebra_table(self, headers, rows, no_wrap_cols=None):
        """
        headers: list of column header strings, in natural RTL reading order
            (e.g. the serial column first)
        rows: list of row value-lists, same column order as headers
        no_wrap_cols: set of 0-based column indices (in the same pre-reversal
            order as headers) that should be sized to fit their content on one
            line instead of wrapping - typically dates, ids, and other short
            fixed-format values. Build it with no_wrap_indices_for(). Columns
            not listed share the rest of the page width and wrap on spaces
            (fpdf2's default wrapmode).

        fpdf2 lays out table columns left-to-right regardless of cell text
        alignment, so the column order is reversed here before building the
        table: the first logical column ends up the rightmost physical one,
        matching RTL reading order.
        """
        no_wrap_cols = no_wrap_cols or set()
        col_widths = self._column_widths(headers, rows, no_wrap_cols)

        headers = list(reversed(headers))
        col_widths = list(reversed(col_widths))
        rows = [list(reversed(row)) for row in rows]

        headings_style = FontFace(family="Amiri", emphasis="", fill_color=HEADER_FILL)
        with self.table(
            text_align="R",
            headings_style=headings_style,
            col_widths=tuple(col_widths),
        ) as table:
            header_row = table.row()
            for h in headers:
                header_row.cell(h)
            for idx, row_vals in enumerate(rows):
                fill = ZEBRA_FILL if idx % 2 == 1 else None
                style = FontFace(fill_color=fill) if fill else None
                row = table.row(style=style)
                for val in row_vals:
                    row.cell(str(val) if val is not None else "")
