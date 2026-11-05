"""Assemble the full book (Word + PDF) from the 18 chapter files.

Usage: python3 bookgen.py
Inputs : ../out/chapter/Ch01..Ch18_Book.docx (built by build.py from content/)
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

N_CH = 18
PARTS = [
    ("الجزء الأول", "صياغة القرار وتقييم شواهده", "Framing the Decision and Assessing Evidence", [1, 2, 3, 4, 5],
     "يبدأ من القرار لا من البيانات: يعرّف القرار المبني على البيانات وخطواته الست، ثم يصوغ القرار، ويجمع شواهده ويقيّم جودتها، ويستدل منها إحصائيا وسببيا."),
    ("الجزء الثاني", "مقارنة البدائل تحت المخاطر", "Comparing Alternatives under Risk", [6, 7, 8, 9, 10],
     "ينتقل من التنبؤ وتقييم النماذج بمنظور القرار إلى تقييم الخيارات: القرار تحت عدم اليقين، والتحسين والمحاكاة، والمفاضلة متعددة المعايير."),
    ("الجزء الثالث", "إيصال القرار وتنفيذه", "Communicating and Implementing the Decision", [11, 12, 13, 14],
     "يراجع الحكم البشري من التحيّزات والضجيج، ويصمم القرار المشترك مع الذكاء الاصطناعي، ويوصل القرار بأمانة، ثم ينفذه ويقيس أثره ويحكمه."),
    ("الجزء الرابع", "القرار في المنظمة والمجتمع", "Decisions in the Organization and Society", [15, 16, 17, 18],
     "يتسع إلى أخلاقيات القرار الخوارزمي والمساءلة، والمنظمة المبنية على البيانات، وتطبيقات من السياق السعودي، ثم يجمع الخطوات الست في قرار متكامل."),
]
APPENDICES = ["ملحق أ: المصطلحات الموحدة", "ملحق ب: المراجع الموحدة", "ملحق ج: فهرس الموضوعات",
              "ملحق د: بنك الأسئلة", "ملحق هـ: دليل المدرس"]
FRONT_LISTS = ["قائمة الأشكال", "قائمة الجداول"]
STEPS = {1: "كلها", 2: "1", 3: "2", 4: "2–3", 5: "3", 6: "3", 7: "3–4", 8: "4", 9: "4", 10: "4",
         11: "5", 12: "5", 13: "5", 14: "6", 15: "كلها", 16: "كلها", 17: "كلها", 18: "1–6"}


def chapter_path(n):
    return os.path.join(ROOT, "out", "chapter", f"Ch{n:02d}_Book.docx")


def chapter_titles():
    t = {}
    for n in range(1, N_CH + 1):
        d = docx.Document(chapter_path(n))
        t[n] = d.paragraphs[1].text.strip()
    return t


# --------------------------------------------------------------------------- glossary and references
def collect_terms_refs():
    terms, refs = {}, {}
    for n in range(1, N_CH + 1):
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
    f.center("Data-Driven Decision Making", size=32, color=AMBER, after=2000)
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
        "وتنتظم الفصول الثمانية عشر في أربعة أجزاء: صياغة القرار وتقييم شواهده (الفصول 1–5)، ومقارنة البدائل تحت المخاطر (6–10)، وإيصال القرار وتنفيذه (11–14)، والقرار في المنظمة والمجتمع (15–18). ويتناول الجزء الأخير ما يحيط بالخطوات كلها: أخلاقيات القرار الخوارزمي والمساءلة في ضوء الأطر الوطنية، وبناء المنظمة المبنية على البيانات، وتطبيقات من السياق السعودي.",
        "ويخدم الكتاب طلاب الماجستير والدكتوراه معا. فمتن الفصول مشترك، وتُختم كل فصل بإطار «للدراسات المتقدمة» يعرض الجدل البحثي في موضوعه، وأسئلة بحثية مفتوحة، وقراءات محكّمة، وتكليفا خاصا بطلاب الدكتوراه، ومخرج تعلم إضافيا مرتبطا بمجالات إطار المؤهلات الوطني. وتُعلَّم الأقسام الاختيارية بوسم «للاستزادة»، وتُجمع الموضوعات الفنية الأعمق في ملحق كل فصل.",
        "والتزم الكتاب ألا يورد حالة أو رقما منسوبا إلا بمصدره، وأن يميز الأرقام التدريبية من الأرقام المنشورة، وأن يوحد مصطلحاته العربية ويذكر المقابل الإنجليزي عند أول وروده فقط. وتُرحَّب ملاحظات الزملاء والطلاب لتحسين الطبعات القادمة.",
    ]:
        f.para(t)

    f.h1("كيف تستعمل هذا الكتاب", page_break=True)
    f.para("يتبع كل فصل بنية ثابتة تساعد الطالب على الدراسة الذاتية والمدرس على التخطيط:", keep=True)
    f.items([
        "سؤال الفصل وحالة افتتاحية حقيقية تطرح القرار قبل الأدوات.",
        "مخرجات تعلم بأفعال قابلة للقياس مرتبطة بمجالات إطار المؤهلات الوطني، وخريطة للفصل بأربعة محاور ومخرج واحد.",
        "أقسام مرقمة، لكل منها شرح وجدول مرقم، وخلاصة في سطر، ومثال محلول في إطار مميز.",
        "حالتان تطبيقيتان ومقارنة بينهما، ثم أدوات تنقل التحليل إلى القرار.",
        "بطاقة الفصل ومعايير تقييمها الخمسة، وهي المخرج الذي يُبنى عليه ما بعده.",
        "إطار «للدراسات المتقدمة» لطلاب الدكتوراه: جدل بحثي وأسئلة مفتوحة وقراءات محكّمة وتكليف.",
        "تطبيق في Excel بأوراق شرح وتمارين محسوبة بالصيغ، وحالة للنقاش ومهمة بحثية.",
        "خلاصة وأسئلة مراجعة بمستويات الفهم والتطبيق والحكم، ومصطلحات ومراجع، وملحق للموضوعات الفنية، وإجابات نموذجية.",
    ])
    f.para("وتعني الألوان في الكتاب ما يلي: الإطار الكريمي خلاصة القسم أو سؤال يُطرح، والإطار الأخضر المزرق مثال محلول، والإطار البنفسجي مادة الدراسات المتقدمة، والوسم الذهبي «للاستزادة» قسم اختياري.")
    f.h2("خريطة المقرر")
    f.table(["الفصل", "العنوان", "الخطوة"],
            [[str(n), titles[n], STEPS[n]] for n in range(1, N_CH + 1)], [1, 5, 1.2])

    # table of contents
    f.h1("المحتويات", page_break=True)
    rows = []
    for key in FRONT_LISTS:
        rows.append([key, str(toc_pages.get(key, "")), {"bold": True, "color": TEAL, "size": 24}])
    for name, sub, _, chs, _ in PARTS:
        rows.append([f"{name}: {sub}", str(toc_pages.get(name, "")), {"bold": True, "color": TEAL, "size": 26}])
        for n in chs:
            rows.append([f"الفصل {ORD[n]}: {titles[n]}", str(toc_pages.get(n, "")), {}])
    for key in APPENDICES:
        rows.append([key, str(toc_pages.get(key, "")), {"bold": True, "color": TEAL, "size": 26}])
    f._plain_table(rows, [8, 1], False)

    # lists of figures and tables (page numbers from the previous PDF pass)
    f.h1("قائمة الأشكال", page_break=True)
    f._plain_table([[f"الشكل {n}\u200e-1: {titles[n]}", str(toc_pages.get(("fig", n), "")), {"size": 22}]
                    for n in range(1, N_CH + 1)], [8, 1], False)
    f.h1("قائمة الجداول", page_break=True)
    caps = toc_pages.get("captions", [])
    f._plain_table([[c, str(pg), {"size": 20}] for c, pg in caps] or [["تُستكمل أرقام الصفحات في التمرير الثاني", "", {}]],
                   [8, 1], False)
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
    if ppr.find(W + "pStyle") is None:
        ppr.insert(0, parse_xml(f'<w:pStyle {W_NS} w:val="BookChapter"/>'))
    return d


def load_chapters():
    import importlib
    sys.path.insert(0, os.path.join(HERE, "content"))
    from textterms import walk as walk_terms
    from advanced import ADV
    chs = {}
    for n in range(1, N_CH + 1):
        ch = walk_terms(importlib.import_module(f"ch{n:02d}").CH)
        ch = dict(ch, advanced=ADV.get(n))
        chs[n] = ch
    return chs


def _texts(o, out):
    if isinstance(o, str):
        out.append(o)
    elif isinstance(o, dict):
        for k, v in o.items():
            if k not in ("en", "lead_en"):
                _texts(v, out)
    elif isinstance(o, (list, tuple)):
        for v in o:
            _texts(v, out)
    return out


def subject_index(chs, terms):
    """Glossary terms -> section numbers where they occur (chapter.section)."""
    secs = []
    for n, ch in chs.items():
        for i, sec in enumerate(ch["sections"], 1):
            secs.append((f"{n}.{i}", " ".join(_texts(sec, []))))
    rows = []
    for ar, (en, _) in terms.items():
        hits = [num for num, txt in secs if ar in txt]
        if hits:
            shown = hits[:10]
            rows.append([ar, en, "، ".join(shown) + (" ..." if len(hits) > 10 else "")])
    return sorted(rows, key=lambda r: ar_key(r[0]))


def _appendix_title(b, text):
    b.page_break()
    b.p(_runs(text, bold=True, color=TEAL, size=36),
        '<w:pStyle w:val="BookPart"/><w:bidi/><w:spacing w:after="160" w:before="320"/>')


def appendices(doc, terms, refs, chs):
    b = Front.__new__(Front)
    b.doc = doc
    b.ch = {"num": 0}
    b.tbl_n = 0
    b.body = doc.element.body
    b.sect = b.body[-1]
    b.numbering = doc.part.numbering_part.element
    b._next_num = 900

    _appendix_title(b, APPENDICES[0])
    b.para("المصطلحات الواردة في جداول المصطلحات بالفصول، مرتبة ترتيبا ألفبائيا، مع أرقام الفصول التي وردت فيها.")
    rows = sorted(([ar, en, "، ".join(str(x) for x in sorted(cs))] for ar, (en, cs) in terms.items()), key=lambda r: ar_key(r[0]))
    b.table(["المصطلح", "المقابل الإنجليزي", "الفصول"], rows, [3, 3, 1])

    _appendix_title(b, APPENDICES[1])
    b.para("المراجع التي استُشهد بها في الفصول، مرتبة ترتيبا ألفبائيا، مع أرقام الفصول التي استُشهد بها فيها.")
    latin = sorted((r for r in refs.values() if re.match(r"[A-Za-z]", r[0])), key=lambda r: r[0].lower())
    arabic = sorted((r for r in refs.values() if not re.match(r"[A-Za-z]", r[0])), key=lambda r: ar_key(r[0]))
    for ref, cs in arabic + latin:
        b.p(_runs(ref, size=21) + _runs(f"   [الفصول: {'، '.join(str(x) for x in sorted(cs))}]", size=19, color=GREY),
            '<w:bidi/><w:spacing w:after="100"/><w:ind w:hanging="360" w:right="360"/><w:jc w:val="start"/>')

    _appendix_title(b, APPENDICES[2])
    b.para("يحيل الفهرس كل مصطلح إلى أرقام الأقسام التي ورد فيها (رقم الفصل ثم رقم القسم)، فيمكن الوصول إليه من رأس القسم في الفصل. ويُكتفى بأول عشرة مواضع للمصطلحات الشائعة.")
    b.table(["الموضوع", "المقابل الإنجليزي", "الأقسام"], subject_index(chs, terms), [2.6, 2.6, 2.4])

    _appendix_title(b, APPENDICES[3])
    b.para("يجمع البنك أسئلة المراجعة بمستوياتها الثلاثة (فهم، وتطبيق، وحكم) مع مفاتيح إجاباتها، ثم أسئلة النقاش، ثم الأسئلة البحثية المفتوحة لطلاب الدكتوراه. ويمكن للمدرس أن يبني منه اختبارات قصيرة أو أسئلة نقاش أو موضوعات مقالات.")
    lv_total = {}
    for n, ch in chs.items():
        b.h2(f"الفصل {ORD[n]}: {ch['title']}")
        qa = [[lvl, q, a] for (lvl, q), a in zip(ch["review"], ch["answers"])]
        for lvl, _, _ in qa:
            lv_total[lvl] = lv_total.get(lvl, 0) + 1
        b.table(["المستوى", "السؤال", "مفتاح الإجابة"], qa, [0.8, 3.2, 3.6])
        b.label("أسئلة النقاش")
        b.items(ch["discussion"])
        if ch.get("advanced"):
            b.label("أسئلة بحثية للدكتوراه")
            b.items(ch["advanced"]["questions"])
    b.para("مجموع أسئلة المراجعة: " + "، ".join(f"{k} {v}" for k, v in lv_total.items()) + ".")

    _appendix_title(b, APPENDICES[4])
    b.h2("فلسفة المقرر وتقييمه")
    for t in [
        "يُدرَّس المقرر بالقرار لا بالأداة: يبدأ كل لقاء بالحالة الافتتاحية وسؤالها، ويُستدعى التحليل لخدمة قرار يتخذه الطلاب ويدافعون عنه. ويُقيَّم الطالب بمخرجات الفصول (البطاقات وتمارين Excel) والمشروع الختامي، دون اختبار نهائي تقليدي.",
        "ويُوزَّع التقييم المقترح على النحو الآتي لطلاب الماجستير: بطاقات الفصول 40%، وتمارين Excel 15%، والمشاركة في نقاش الحالات 10%، والمشروع الختامي 35%. ولطلاب الدكتوراه: بطاقات الفصول 30%، وتكليفات «للدراسات المتقدمة» 25%، والمشاركة 10%، والمشروع الختامي بمقترح بحثي 35%.",
        "وتُقيَّم كل بطاقة بمعاييرها الخمسة (صفر، واحد، اثنان)، ويُعاد للطالب تقييم مكتوب قبل الفصل التالي لأن البطاقات تتراكم: فكل فصل يبني على بطاقة ما قبله، ويجمعها ملف القرار في الفصل الأخير.",
    ]:
        b.para(t)
    b.h2("الخطة الأسبوعية المقترحة")
    plan = [("1", "1"), ("2", "2"), ("3", "3"), ("4", "4"), ("5", "5"), ("6", "6"), ("7", "7"), ("8", "8"),
            ("9", "9"), ("10", "10"), ("11", "11 و12"), ("12", "13"), ("13", "14"), ("14", "15 و16"), ("15", "17 و18 وعرض المشروعات")]
    rows = []
    for wk, cs in plan:
        nums = [int(x) for x in re.findall(r"\d+", cs)]
        files = "، ".join(chs[x]["app"]["file"] for x in nums if x in chs)
        rows.append([wk, cs, " / ".join(chs[x]["title"] for x in nums if x in chs), files])
    b.table(["الأسبوع", "الفصل", "الموضوع", "ملف Excel"], rows, [0.8, 1.2, 3.4, 2.8])
    b.para("ويمكن في مسار القرار التنظيمي التوسع في الفصول 12 و15 و16، وفي مسار القرار التشغيلي التوسع في الفصول 6 و9 و14، مع إسناد الأقسام الموسومة «للاستزادة» للقراءة الذاتية.")
    b.h2("دليل الفصول")
    rows = []
    for n, ch in chs.items():
        cases = "؛ ".join(s["t"].split(":", 1)[-1].strip() for s in ch["sections"] if s["axis"] == 1 and not s["t"].startswith("مقارنة"))
        rows.append([str(n), cases, f"{ch['app']['file']} ({ch['app']['time']})", ch["output"]])
    b.table(["الفصل", "الحالتان التطبيقيتان", "التطبيق والزمن", "المخرج المطلوب"], rows, [0.6, 2.6, 2.0, 3.2])
    b.h2("إرشادات التدريس")
    b.items([
        "ابدأ بالحالة الافتتاحية واطلب من الطلاب قرارا مكتوبا قبل الشرح، ثم عُد إليه في نهاية اللقاء لمقارنة القرار قبل الأدوات وبعدها.",
        "استعمل المثال المحلول في كل قسم تمرينا صفيا قصيرا، ثم اطلب تطبيقه على مؤسسة الطالب.",
        "خصص ربع اللقاء لتمرين Excel في مجموعات، واطلب من كل مجموعة تفسير النتيجة بلغة القرار لا بلغة الأرقام.",
        "اعرض الأخطاء المنهجية في نهاية كل فصل حالات قصيرة يكتشف الطلاب الخطأ فيها.",
        "لطلاب الدكتوراه: خصص لقاء شهريا لعرض قراءة محكّمة ونقدها، ومتابعة تكليفات «للدراسات المتقدمة».",
        "وثّق مصدر كل رقم تستعمله في الصف، وميّز الرقم التدريبي من المنشور، كما يفعل الكتاب.",
    ])


def assemble(path, titles, toc_pages, terms, refs, chs):
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
    appendices(comp.doc, terms, refs, chs)
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
    found = {"captions": []}

    def walk(items):
        for it in items:
            if isinstance(it, list):
                walk(it)
                continue
            t = (it.title or "").strip()
            pg = r.get_destination_page_number(it) + 1
            if t.startswith("جدول "):
                found["captions"].append((t.replace("\u200e", "\u200e"), pg))
                continue
            for n, tt in titles.items():
                if t == tt and n not in found:
                    found[n] = pg
                    found[("fig", n)] = pg
            for name, sub, *_ in PARTS:
                if t == sub and name not in found:
                    found[name] = pg
            for key in APPENDICES + FRONT_LISTS:
                if t == key and key not in found:
                    found[key] = pg
    walk(r.outline)
    return found, len(r.pages)


def strip_caption_bookmarks(pdf_path):
    """Rebuild the PDF outline without the (many) table-caption entries."""
    from pypdf import PdfWriter
    r = PdfReader(pdf_path)
    w = PdfWriter(clone_from=r)
    w._root_object.pop("/Outlines", None)

    def add(items, parent=None):
        last = None
        for it in items:
            if isinstance(it, list):
                if last is not None:
                    add(it, last)
                continue
            t = (it.title or "").strip()
            if t.startswith("جدول "):
                continue
            last = w.add_outline_item(t, r.get_destination_page_number(it), parent=parent)
    add(r.outline)
    w.page_mode = "/UseOutlines"
    with open(pdf_path, "wb") as fh:
        w.write(fh)


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
    titles = chapter_titles()
    terms, refs = collect_terms_refs()
    chs = load_chapters()
    path = os.path.join(OUT, "Data_Driven_Decision_Making_Book.docx")
    pages = {}
    for attempt in range(4):
        assemble(path, titles, pages, terms, refs, chs)
        pdf = to_pdf(path)
        new, total = outline_pages(pdf, titles)
        print("pass", attempt + 1, "pages", total, "found", len(new), "captions", len(new["captions"]))
        if new == pages:
            break
        pages = new
    strip_caption_bookmarks(pdf)
    print({k: v for k, v in pages.items() if k != "captions"})
    print("terms", len(terms), "refs", len(refs))


if __name__ == "__main__":
    main()
