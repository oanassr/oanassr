"""Assemble the full book (Word + PDF) from the 13 chapter files.

Usage: python3 bookgen.py
Inputs : ../src_chapters/Ch01..Ch09_Book.docx (from Drive), ../out/chapter/Ch10..Ch13_Book.docx
Outputs: ../out/book/Data_Driven_Decision_Making_Book.docx and ../out/book/Data_Driven_Decision_Making_Book.pdf

The table of contents is static: page numbers are read from a first PDF pass
(LibreOffice bookmarks) and written into the second pass.
"""
import copy
import os
import re
import subprocess
import sys

import docx
from docx.oxml import parse_xml
from docx.shared import Inches
from docxcompose.composer import Composer
from pypdf import PdfReader

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docgen import DocBuilder, W_NS, _runs, TEAL, AMBER, CREAM, GREY  # noqa: E402
from textutil import ORD  # noqa: E402
from normalize import normalize_document  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, "..")
OUT = os.path.join(ROOT, "out", "book")
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

PARTS = [
    ("الجزء الأول", "تأطير القرار وتقييم شواهده", "Framing the Decision and Assessing Evidence", [1, 2, 3, 4],
     "يبدأ من القرار لا من البيانات: يعرّف القرار المبني على البيانات وخطواته الست، ثم يصوغ القرار، ويجمع شواهده ويقيّم جودتها، ويستدل منها إحصائيا."),
    ("الجزء الثاني", "مقارنة البدائل تحت المخاطر", "Comparing Alternatives under Risk", [5, 6, 7, 8, 9],
     "ينتقل من التنبؤ وتقييم النماذج بمنظور القرار إلى تقييم الخيارات: القرار تحت عدم اليقين، والتحسين والمحاكاة، والمفاضلة متعددة المعايير."),
    ("الجزء الثالث", "إيصال القرار ومتابعته", "Communicating and Following Through", [10, 11, 12, 13],
     "يراجع الحكم البشري من التحيّزات والضجيج، ويوصل القرار بأمانة، وينفذه ويقيس أثره ويحكمه، ثم يجمع الخطوات الست في قرار متكامل."),
]


def chapter_path(n):
    if n <= 9:
        return os.path.join(ROOT, "src_chapters", f"Ch{n:02d}_Book.docx")
    return os.path.join(ROOT, "out", "chapter", f"Ch{n:02d}_Book.docx")


def chapter_titles():
    t = {}
    for n in range(1, 14):
        d = docx.Document(chapter_path(n))
        t[n] = d.paragraphs[1].text.strip()
    return t


# --------------------------------------------------------------------------- glossary and references
def collect_terms_refs():
    terms, refs = {}, {}
    for n in range(1, 14):
        d = docx.Document(chapter_path(n))
        body = list(d.element.body)
        mode = None
        for el in body:
            tag = el.tag.replace(W, "")
            txt = "".join(t.text or "" for t in el.iter(W + "t")).strip() if tag == "p" else ""
            if tag == "p" and el.find(f".//{W}pStyle") is not None and el.find(f".//{W}pStyle").get(W + "val") == "Heading1":
                mode = {"المصطلحات": "terms", "المراجع": "refs"}.get(txt)
                continue
            if mode == "terms" and tag == "tbl":
                for tr in el.findall(f"{W}tr")[1:]:
                    cells = ["".join(t.text or "" for t in tc.iter(W + "t")).strip() for tc in tr.findall(f"{W}tc")]
                    if len(cells) >= 2 and cells[0]:
                        terms.setdefault(cells[0], (cells[1], set()))[1].add(n)
                mode = None
            elif mode == "refs" and tag == "p" and txt:
                ref = txt.split(" — ")[0].strip()
                key = re.sub(r"[^A-Za-z0-9؀-ۿ]", "", ref)[:60].lower()
                refs.setdefault(key, [ref, set()])[1].add(n)
            elif mode == "refs" and tag == "p" and el.find(f".//{W}br") is not None:
                mode = None
    return terms, refs


def ar_key(s):
    s = re.sub(r"^ال", "", s.strip())
    return s.translate(str.maketrans("أإآة", "اااه"))


# --------------------------------------------------------------------------- front matter builder
class Front(DocBuilder):
    def __init__(self):
        super().__init__({"num": 0, "title": "إدارة اتخاذ القرار على أساس البيانات"})

    def center(self, text, size=26, color=None, bold=False, before=0, after=120):
        self.p(_runs(text, size=size, color=color, bold=bold),
               f'<w:bidi/><w:spacing w:before="{before}" w:after="{after}"/><w:jc w:val="center"/>')

    def cover(self, img):
        par = self.doc.add_paragraph()
        par.add_run().add_picture(img, width=Inches(8.26), height=Inches(11.68))
        ppr = par._p.get_or_add_pPr()
        ppr.append(parse_xml(f'<w:spacing {W_NS} w:before="0" w:after="0"/>'))
        ppr.append(parse_xml(f'<w:jc {W_NS} w:val="center"/>'))
        # the cover paragraph closes a section with zero margins and no footer
        ppr.append(parse_xml(
            f'<w:sectPr {W_NS}><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>'))

    def table_rows(self, rows, widths, header=None, bold_first=False):
        self.table(header or ["", ""], rows, widths) if header else self._plain_table(rows, widths, bold_first)

    def _plain_table(self, rows, widths, bold_first):
        total = 9300
        s = sum(widths)
        ws = [int(w * total / s) for w in widths]
        grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in ws)
        out = [f'<w:tbl {W_NS}><w:tblPr><w:bidiVisual/><w:tblW w:type="dxa" w:w="{total}"/>'
               '<w:tblBorders><w:insideH w:val="single" w:color="DDD6C6" w:sz="4"/></w:tblBorders></w:tblPr>'
               f'<w:tblGrid>{grid}</w:tblGrid>']
        for r in rows:
            cells = []
            for j, (c, w) in enumerate(zip(r, ws)):
                style = r[-1] if isinstance(r[-1], dict) else {}
                if isinstance(c, dict):
                    continue
                bold = style.get("bold", False) or (bold_first and j == 0)
                color = style.get("color")
                size = style.get("size", 24)
                jc = "end" if j == len(ws) - 1 else "start"
                cells.append(f'<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="{w}"/></w:tcPr><w:p><w:pPr><w:bidi/>'
                             f'<w:spacing w:before="40" w:after="40"/><w:jc w:val="{jc}"/></w:pPr>'
                             f'{_runs(str(c), size=size, bold=bold, color=color)}</w:p></w:tc>')
            out.append("<w:tr><w:trPr><w:cantSplit/></w:trPr>" + "".join(cells) + "</w:tr>")
        out.append("</w:tbl>")
        self._add("".join(out))


def front_matter(path, titles, toc_pages):
    f = Front()
    f.cover(os.path.join(ROOT, "img", "book_cover.jpg"))

    # title page
    f.center("", after=1800)
    f.center("إدارة اتخاذ القرار على أساس البيانات", size=56, color=TEAL, bold=True, after=200)
    f.center("Data-Driven Decision Making", size=32, color=AMBER, after=600)
    f.center("كتاب المقرر", size=30, bold=True, after=1400)
    f.center("تأليف", size=24, color=GREY, after=80)
    f.center("د. عثمان نصر", size=36, color=TEAL, bold=True, after=120)
    f.center("أستاذ مساعد – قسم المعلوماتية للأعمال", size=24, color=GREY, after=40)
    f.center("كلية الأعمال – جامعة الملك خالد", size=24, color=GREY, after=600)
    f.center("الطبعة الأولى (تجريبية) – 1448هـ / 2026م", size=24, color=GREY, before=600)
    f.page_break()

    # rights page
    f.center("", after=10500)
    for t in ["جميع الحقوق محفوظة للمؤلف © 2026.",
              "أُعدّ هذا الكتاب لتدريس مقرر إدارة اتخاذ القرار على أساس البيانات في الدراسات العليا.",
              "الأرقام في الأمثلة المحلولة وتمارين Excel تدريبية ما لم يُنسب الرقم إلى مصدر منشور.",
              "الحالات المنشورة موثقة بمصادرها في نهاية كل فصل وفي ملحق المراجع الموحدة.",
              "يرافق كل فصل ملف عرض تقديمي وملف تمارين Excel."]:
        f.small(t, color=GREY, size=22)
    f.page_break()

    # preface
    f.h1("المقدمة")
    for t in [
        "يقف هذا المقرر بعد مقررات التحليل في الدراسات العليا: فهي تضع الأداة التحليلية في يد الطالب، وهو يعلّمه متى يستعملها، ولماذا، وماذا يقرر بها. ولذلك يقوم الكتاب على مبدأ واحد: لا يُعاد تدريس مهارة تحليلية، بل تُستدعى داخل قرار يتخذه الطالب ويدافع عنه.",
        "ويُبنى الكتاب على ست خطوات لاتخاذ القرار المبني على البيانات: صياغة القرار، وجمع الشواهد وتقييمها، والتحليل والتنبؤ، وتقييم الخيارات، والقرار والتواصل، والتنفيذ والتعلم. ويعرض الفصل الأول الخطوات كاملة، ثم يخصص لكل خطوة فصلا أو أكثر، ويجمعها الفصل الأخير في حالة متكاملة.",
        "وتنتظم الفصول الثلاثة عشر في ثلاثة أجزاء تقابل ثلاث مهارات: تأطير القرار وتقييم شواهده (الفصول 1–4)، ومقارنة البدائل تحت المخاطر (الفصول 5–9)، وإيصال القرار ومتابعته (الفصول 10–13). ولكل مهارة مخرجات تُقيَّم من خلال بطاقات الفصول وتمارينها.",
        "ويخدم الكتاب مسارين معا. فالفصول مشتركة، ويتمايز المساران في زاوية القرار ودور الطالب في الحالات والمشروع الختامي: مسار القرار التنظيمي يقيّم فيه الطالب الجاهزية ويصمم الحوكمة، ومسار القرار التشغيلي يبني فيه الطالب نموذج القرار وتجربته.",
        "والتزم الكتاب ألا يورد حالة أو رقما منسوبا إلا بمصدره، وأن يميز الأرقام التدريبية من الأرقام المنشورة. وتُرحَّب ملاحظات الزملاء والطلاب لتحسين الطبعات القادمة.",
    ]:
        f.para(t)

    f.h1("كيف تستعمل هذا الكتاب", page_break=True)
    f.para("يتبع كل فصل بنية ثابتة تساعد الطالب على الدراسة الذاتية والمدرس على التخطيط:", keep=True)
    f.items([
        "سؤال الفصل وحالة افتتاحية حقيقية تطرح القرار قبل الأدوات.",
        "مخرجات تعلم بأفعال قابلة للقياس، وخريطة للفصل بأربعة محاور ومخرج واحد.",
        "أقسام مرقمة، لكل منها شرح وجدول أو بطاقة، وخلاصة في سطر، ومثال محلول.",
        "حالتان تطبيقيتان ومقارنة بينهما، ثم أدوات تنقل التحليل إلى القرار.",
        "بطاقة الفصل ومعايير تقييمها الخمسة، وهي المخرج الذي يُبنى عليه ما بعده.",
        "تطبيق في Excel بأوراق شرح وتمارين محسوبة بالصيغ، وحالة للنقاش ومهمة بحثية.",
        "خلاصة وأسئلة مراجعة بمستويات الفهم والتطبيق والحكم، ومصطلحات ومراجع، وإجابات نموذجية في ملحق.",
    ])
    f.h2("خريطة المقرر")
    f.table(["الفصل", "العنوان", "الخطوة"],
            [[str(n), titles[n], s] for n, s in zip(range(1, 14),
             ["كلها", "1", "2", "2–3", "3", "3–4", "4", "4", "4", "5", "5", "6", "1–6"])], [1, 5, 1.2])

    # table of contents
    f.h1("المحتويات", page_break=True)
    rows = []
    for name, sub, _, chs, _ in PARTS:
        rows.append([f"{name}: {sub}", str(toc_pages.get(name, "")), {"bold": True, "color": TEAL, "size": 26}])
        for n in chs:
            rows.append([f"الفصل {ORD[n]}: {titles[n]}", str(toc_pages.get(n, "")), {}])
    for key in ["ملحق أ: المصطلحات الموحدة", "ملحق ب: المراجع الموحدة"]:
        rows.append([key, str(toc_pages.get(key, "")), {"bold": True, "color": TEAL, "size": 26}])
    f._plain_table(rows, [8, 1], False)
    f.doc.save(path)


def part_page(doc, part):
    name, sub, en, chs, desc = part
    b = Front.__new__(Front)
    b.doc = doc
    b.body = doc.element.body
    b.sect = b.body[-1]
    b.page_break()
    b.p("", '<w:spacing w:before="2400"/>')
    b.p(_runs(name, size=36, color=AMBER, bold=True), '<w:bidi/><w:spacing w:after="120"/><w:jc w:val="center"/>')
    b.p(_runs(sub, size=60, color=TEAL, bold=True),
        '<w:pStyle w:val="BookPart"/><w:bidi/><w:spacing w:after="160"/><w:jc w:val="center"/>')
    b.p(_runs(en, size=28, color=GREY), '<w:bidi/><w:spacing w:after="600"/><w:jc w:val="center"/>')
    b.p(_runs(desc, size=26), '<w:bidi/><w:spacing w:after="400" w:line="360"/><w:ind w:left="900" w:right="900"/><w:jc w:val="center"/>')
    for n in chs:
        b.p(_runs(f"الفصل {ORD[n]}", size=24, color=GREY), '<w:bidi/><w:spacing w:after="0"/><w:jc w:val="center"/>')


def add_styles(doc):
    styles = doc.styles.element
    for sid, name in (("BookPart", "Book Part"), ("BookChapter", "Book Chapter")):
        if styles.find(f'{W}style[@{W}styleId="{sid}"]') is None:
            styles.append(parse_xml(
                f'<w:style {W_NS} w:type="paragraph" w:customStyle="1" w:styleId="{sid}"><w:name w:val="{name}"/>'
                f'<w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>'))


def clean_chapter(d):
    """Remove comment anchors and mark the chapter title for bookmarks/TOC."""
    body = d.element.body
    for tag in ("commentRangeStart", "commentRangeEnd", "commentReference"):
        for el in body.iter(W + tag):
            el.getparent().remove(el)
    normalize_document(d)
    title = d.paragraphs[1]._p
    ppr = title.get_or_add_pPr()
    ppr.insert(0, parse_xml(f'<w:pStyle {W_NS} w:val="BookChapter"/>'))
    return d


def appendices(doc, terms, refs):
    b = Front.__new__(Front)
    b.doc = doc
    b.body = doc.element.body
    b.sect = b.body[-1]
    b.numbering = doc.part.numbering_part.element
    b._next_num = 900
    b.page_break()
    b.p(_runs("ملحق أ: المصطلحات الموحدة", bold=True, color=TEAL, size=36),
        '<w:pStyle w:val="BookPart"/><w:bidi/><w:spacing w:after="160" w:before="320"/>')
    b.para("المصطلحات الواردة في جداول المصطلحات بالفصول، مرتبة ترتيبا ألفبائيا، مع أرقام الفصول التي وردت فيها.")
    rows = sorted(([ar, en, "، ".join(str(x) for x in sorted(chs))] for ar, (en, chs) in terms.items()), key=lambda r: ar_key(r[0]))
    b.table(["المصطلح", "بالإنجليزية", "الفصول"], rows, [3, 3, 1])
    b.page_break()
    b.p(_runs("ملحق ب: المراجع الموحدة", bold=True, color=TEAL, size=36),
        '<w:pStyle w:val="BookPart"/><w:bidi/><w:spacing w:after="160" w:before="320"/>')
    b.para("المراجع التي استُشهد بها في الفصول، مرتبة ترتيبا ألفبائيا، مع أرقام الفصول التي استُشهد بها فيها.")
    latin = sorted((r for r in refs.values() if re.match(r"[A-Za-z]", r[0])), key=lambda r: r[0].lower())
    arabic = sorted((r for r in refs.values() if not re.match(r"[A-Za-z]", r[0])), key=lambda r: ar_key(r[0]))
    for ref, chs in arabic + latin:
        b.p(_runs(ref, size=21) + _runs(f"   [الفصول: {'، '.join(str(x) for x in sorted(chs))}]", size=19, color=GREY),
            '<w:bidi/><w:spacing w:after="100"/><w:ind w:hanging="360" w:right="360"/><w:jc w:val="start"/>')


def assemble(path, titles, toc_pages, terms, refs):
    front = path.replace(".docx", "_front.docx")
    front_matter(front, titles, toc_pages)
    master = docx.Document(front)
    add_styles(master)
    comp = Composer(master)
    for part in PARTS:
        part_page(comp.doc, part)
        for n in part[3]:
            comp.doc.add_page_break()
            d = clean_chapter(docx.Document(chapter_path(n)))
            add_styles(d)
            comp.append(d)
    appendices(comp.doc, terms, refs)
    s = comp.doc.settings.element
    normalize_document(comp.doc)
    comp.doc.core_properties.title = "إدارة اتخاذ القرار على أساس البيانات"
    comp.doc.core_properties.author = "د. عثمان نصر"
    comp.save(path)
    os.remove(front)


def to_pdf(docx_path):
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", os.path.dirname(docx_path), docx_path],
                   check=True, capture_output=True, timeout=900)
    return docx_path.replace(".docx", ".pdf")


def outline_pages(pdf_path, titles):
    r = PdfReader(pdf_path)
    found = {}

    def walk(items):
        for it in items:
            if isinstance(it, list):
                walk(it)
                continue
            t = (it.title or "").strip()
            pg = r.get_destination_page_number(it) + 1
            for n, tt in titles.items():
                if t == tt and n not in found:
                    found[n] = pg
            for name, sub, *_ in PARTS:
                if t == sub and name not in found:
                    found[name] = pg - 0
            for key in ["ملحق أ: المصطلحات الموحدة", "ملحق ب: المراجع الموحدة"]:
                if t == key:
                    found[key] = pg
    walk(r.outline)
    return found, len(r.pages)


def export_chapters_1_9():
    """Write typographically normalized copies of chapters 1-9 next to chapters 10-13."""
    dst = os.path.join(ROOT, "out", "chapter")
    os.makedirs(dst, exist_ok=True)
    for n in range(1, 10):
        d = docx.Document(chapter_path(n))
        normalize_document(d)
        d.core_properties.author = "د. عثمان نصر"
        d.save(os.path.join(dst, f"Ch{n:02d}_Book.docx"))


def main():
    os.makedirs(OUT, exist_ok=True)
    export_chapters_1_9()
    titles = chapter_titles()
    terms, refs = collect_terms_refs()
    path = os.path.join(OUT, "Data_Driven_Decision_Making_Book.docx")
    pages = {}
    for attempt in range(3):
        assemble(path, titles, pages, terms, refs)
        pdf = to_pdf(path)
        new, total = outline_pages(pdf, titles)
        print("pass", attempt + 1, "pages", total, "found", len(new), "of", 13 + 3 + 2)
        if new == pages:
            break
        pages = new
    print({k: v for k, v in pages.items()})
    print("terms", len(terms), "refs", len(refs))


if __name__ == "__main__":
    main()
