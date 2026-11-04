"""Shared styling for the chapter Excel workbooks (instructions, case walkthrough, exercises)."""
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

TEAL, AMBER, INPUT, FORMULA, CREAM = "0F4C5C", "E09F3E", "FFF2CC", "DDEBF7", "FBF7EF"
thin = Side(style="thin", color="B0B8BA")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True, readingOrder=2)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True, readingOrder=2)


def font(bold=False, color="14213D", size=11):
    return Font(name="Cairo", bold=bold, color=color, size=size)


class Book:
    def __init__(self):
        self.wb = Workbook()
        self.wb.remove(self.wb.active)

    def sheet(self, name, widths=(34, 16, 16, 16, 16, 16)):
        ws = self.wb.create_sheet(name)
        ws.sheet_view.rightToLeft = True
        ws.sheet_view.showGridLines = False
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        return Sheet(ws)

    def save(self, path):
        self.wb.save(path)


class Sheet:
    def __init__(self, ws):
        self.ws = ws
        self.row = 1

    def title(self, text, span=6):
        ws = self.ws
        c = ws.cell(self.row, 1, text)
        c.font = font(True, "FFFFFF", 14)
        c.fill = PatternFill("solid", fgColor=TEAL)
        c.alignment = RIGHT
        ws.merge_cells(start_row=self.row, start_column=1, end_row=self.row, end_column=span)
        ws.row_dimensions[self.row].height = 30
        self.row += 1

    def note(self, text, span=6, color="555555", bold=False, height=None, fill=None):
        ws = self.ws
        c = ws.cell(self.row, 1, text)
        c.font = font(bold, color, 10.5)
        c.alignment = RIGHT
        if fill:
            c.fill = PatternFill("solid", fgColor=fill)
        ws.merge_cells(start_row=self.row, start_column=1, end_row=self.row, end_column=span)
        lines = max(1, len(text) // (span * 13) + 1)
        ws.row_dimensions[self.row].height = height or max(20, 18 * lines)
        self.row += 1

    def section(self, text, span=6):
        ws = self.ws
        c = ws.cell(self.row, 1, text)
        c.font = font(True, TEAL, 12)
        c.alignment = RIGHT
        c.fill = PatternFill("solid", fgColor=CREAM)
        ws.merge_cells(start_row=self.row, start_column=1, end_row=self.row, end_column=span)
        ws.row_dimensions[self.row].height = 24
        self.row += 1

    def header(self, values):
        for j, v in enumerate(values, 1):
            c = self.ws.cell(self.row, j, v)
            c.font = font(True, "FFFFFF", 10.5)
            c.fill = PatternFill("solid", fgColor=TEAL)
            c.alignment = CENTER
            c.border = BORDER
        self.ws.row_dimensions[self.row].height = 34
        self.row += 1
        return self.row - 1

    def put(self, r, c, value, kind="text", fmt=None, bold=False):
        """kind: text | input | formula | label"""
        cell = self.ws.cell(r, c, value)
        cell.border = BORDER
        cell.alignment = CENTER if c > 1 else RIGHT
        cell.font = font(bold or kind == "label")
        if kind == "input":
            cell.fill = PatternFill("solid", fgColor=INPUT)
        elif kind == "formula":
            cell.fill = PatternFill("solid", fgColor=FORMULA)
        if fmt:
            cell.number_format = fmt
        return cell

    def line(self, values, kinds=None, fmts=None, height=None):
        r = self.row
        for j, v in enumerate(values, 1):
            if v is None:
                continue
            k = kinds[j - 1] if kinds else ("label" if j == 1 else "text")
            f = fmts[j - 1] if fmts else None
            self.put(r, j, v, k, f)
        if height:
            self.ws.row_dimensions[r].height = height
        self.row += 1
        return r

    def skip(self, n=1):
        self.row += n

    def dropdown(self, cell_ref, options):
        dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=True)
        self.ws.add_data_validation(dv)
        dv.add(cell_ref)


def instructions(book, ch, sheets_desc):
    s = book.sheet("التعليمات", (100,))
    s.title(f"تمرين الفصل {ch['num']}: {ch['title']} ({ch['app']['file']})", span=1)
    s.note("الأوراق:", span=1, bold=True, color="14213D")
    for i, d in enumerate(sheets_desc, 1):
        s.note(f"{i}) {d}", span=1, color="14213D")
    s.note("الأعداد في التمرين تدريبية لأغراض التدريب.", span=1)
    s.note("الخلايا الصفراء مدخلات الطالب، والخلايا الزرقاء صيغ فلا تعدّلها.", span=1)
    c = s.ws.cell(s.row, 1, "مدخل")
    c.fill = PatternFill("solid", fgColor=INPUT)
    c.font = font()
    c = s.ws.cell(s.row + 1, 1, "صيغة")
    c.fill = PatternFill("solid", fgColor=FORMULA)
    c.font = font()
    return s


def walkthrough(book, ch, blocks):
    """blocks: list of (heading, [lines])"""
    s = book.sheet("شرح_الحالة", (110,))
    s.title("شرح الحالة خطوة بخطوة", span=1)
    for k, (head, lines) in enumerate(blocks, 1):
        s.section(f"{k}. {head}", span=1)
        for ln in lines:
            s.note(ln, span=1, color="14213D")
    return s


def card_sheet(book, name, title, fields, criteria):
    """Generic decision-card + self-assessment sheet (0/1/2 per criterion, total out of 2n)."""
    s = book.sheet(name, (44, 60, 14))
    s.title(title, span=3)
    for f in fields:
        r = s.line([f, ""], ["label", "input"], height=42)
        s.ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    s.skip()
    s.header(["المعيار", "درجة التقييم الذاتي (0 أو 1 أو 2)", "الدرجة"])
    first = s.row
    for c in criteria:
        r = s.line([c, "", None], ["label", "input", "formula"])
        s.put(r, 3, f'=IF(B{r}="",0,MIN(2,MAX(0,B{r})))', "formula")
        s.dropdown(f"B{r}", ["0", "1", "2"])
    last = s.row - 1
    tot = s.line([f"المجموع من {2 * len(criteria)}", None, None], ["label", "text", "formula"])
    s.put(tot, 3, f"=SUM(C{first}:C{last})", "formula")
    hi, mid = round(0.8 * 2 * len(criteria)), round(0.5 * 2 * len(criteria))
    r = s.line(["المستوى", None, None])
    s.put(r, 2, f'=IF(COUNTBLANK(B{first}:B{last})>0,"أكمل التقييم",IF(C{tot}>={hi},"بطاقة قوية",IF(C{tot}>={mid},"بطاقة تحتاج إلى تحسين","بطاقة ضعيفة")))', "formula")
    s.line(["المعيار الأضعف", ""], ["label", "input"])
    s.line(["ما سأفعله لتحسينه", ""], ["label", "input"])
    return s
