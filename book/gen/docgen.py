"""Render a chapter (content dict) to a Word book chapter matching the Ch01–Ch09 design."""
import copy
import os
from xml.sax.saxutils import escape

import docx
from docx.oxml import parse_xml
from docx.shared import Inches

from textutil import split_runs, ORD
from textterms import dedupe_english, walk as walk_terms

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "..", "templates", "Ch09_Book.docx")

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
TEAL, AMBER, CREAM, GREY = "0F4C5C", "E09F3E", "FFF4DD", "555555"
EX_FILL, EX_BAR = "EAF2F3", "2C7A7B"        # worked example box
ADV_FILL, ADV_BAR = "F2EDF5", "6B4E71"      # advanced (doctoral) box
FONT = '<w:rFonts w:ascii="Cairo" w:cs="Cairo" w:eastAsia="Cairo" w:hAnsi="Cairo"/>'
# running text: Naskh for Arabic (complex script), Cairo for Latin
TEXT_FONT = '<w:rFonts w:ascii="Cairo" w:cs="Noto Naskh Arabic" w:eastAsia="Cairo" w:hAnsi="Cairo"/>'
LINE = "360"


def _rpr(bold=False, color=None, size=26, latin=False, explicit_nobold=False, body=False):
    b = "<w:b/><w:bCs/>" if bold else ('<w:b w:val="false"/><w:bCs w:val="false"/>' if explicit_nobold else "")
    c = f'<w:color w:val="{color}"/>' if color else ""
    rtl = '<w:rtl w:val="false"/>' if latin else "<w:rtl/>"
    f = TEXT_FONT if body else FONT
    return f"<w:rPr>{f}{b}{c}<w:sz w:val=\"{size}\"/><w:szCs w:val=\"{size}\"/>{rtl}</w:rPr>"


def _runs(text, **kw):
    out = []
    for seg, latin in split_runs(text):
        out.append(f'<w:r>{_rpr(latin=latin, **kw)}<w:t xml:space="preserve">{escape(seg)}</w:t></w:r>')
    return "".join(out)


W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def ensure_styles(doc):
    """Custom styles: chapter title / part title (TOC + running header) and table captions."""
    styles = doc.styles.element
    for sid, name, lvl in (("BookPart", "Book Part", 0), ("BookChapter", "Book Chapter", 0), ("TableCaption", "Table Caption", 2)):
        if styles.find(f'{W}style[@{W}styleId="{sid}"]') is None:
            styles.append(parse_xml(
                f'<w:style {W_NS} w:type="paragraph" w:customStyle="1" w:styleId="{sid}"><w:name w:val="{name}"/>'
                f'<w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:outlineLvl w:val="{lvl}"/></w:pPr></w:style>'))


def add_running_header(doc):
    """Header with the current chapter title (STYLEREF to the Book Chapter style) and a thin rule."""
    sec = doc.sections[0]
    hdr = sec.header
    hdr.is_linked_to_previous = False
    p = hdr.paragraphs[0]._p
    for ch in list(p):
        p.remove(ch)
    rpr = (f'<w:rPr>{FONT}<w:color w:val="8A949B"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rtl/></w:rPr>')
    p.append(parse_xml(f'<w:pPr {W_NS}><w:pBdr><w:bottom w:val="single" w:color="D6CFC0" w:sz="4" w:space="4"/></w:pBdr>'
                       '<w:bidi/><w:jc w:val="start"/></w:pPr>'))
    for xml in (f'<w:r {W_NS}>{rpr}<w:fldChar w:fldCharType="begin"/></w:r>',
                f'<w:r {W_NS}>{rpr}<w:instrText xml:space="preserve"> STYLEREF "Book Chapter" </w:instrText></w:r>',
                f'<w:r {W_NS}>{rpr}<w:fldChar w:fldCharType="separate"/></w:r>',
                f'<w:r {W_NS}>{rpr}<w:t> </w:t></w:r>',
                f'<w:r {W_NS}>{rpr}<w:fldChar w:fldCharType="end"/></w:r>'):
        p.append(parse_xml(xml))


NQF = [("المعرفة والفهم", ("يعرّف", "يميّز", "يميز", "يفسّر", "يفسر", "يشرح", "يصف", "يحدد")),
       ("القيم والاستقلالية والمسؤولية", ("يوثّق", "يوثق", "يدافع", "يلتزم", "يراجع")),
       ("المهارات", ())]


def nqf_domain(verb):
    for dom, verbs in NQF:
        if verb in verbs:
            return dom
    return "المهارات"


class DocBuilder:
    def __init__(self, ch):
        self.ch = ch
        self.doc = docx.Document(TEMPLATE)
        body = self.doc.element.body
        for el in list(body):
            if not el.tag.endswith("sectPr"):
                body.remove(el)
        # drop the template's picture and comments so nothing of Ch09 remains
        part = self.doc.part
        for rid, rel in list(part.rels.items()):
            if rel.reltype.endswith("/image"):
                part.drop_rel(rid)
            if rel.reltype.endswith("/comments"):
                part.drop_rel(rid)
        self.body = body
        self.sect = body[-1]
        self.numbering = part.numbering_part.element
        self._next_num = 100
        self.tbl_n = 0
        ensure_styles(self.doc)
        add_running_header(self.doc)

    # ---------- low level ----------
    def _add(self, xml):
        el = parse_xml(xml)
        self.sect.addprevious(el)
        return el

    def p(self, inner, ppr=""):
        return self._add(f"<w:p {W_NS}><w:pPr>{ppr}</w:pPr>{inner}</w:p>")

    def new_num(self, abstract=3):
        nid = self._next_num
        self._next_num += 1
        self.numbering.append(parse_xml(
            f'<w:num {W_NS} w:numId="{nid}"><w:abstractNumId w:val="{abstract}"/>'
            f'<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>'))
        return nid

    # ---------- blocks ----------
    def kicker(self, text):
        self.p(_runs(text, bold=True, color=AMBER, size=26), '<w:bidi/><w:spacing w:after="80"/>')

    def title(self, text):
        self.p(_runs(text, bold=True, color=TEAL, size=52), '<w:pStyle w:val="BookChapter"/><w:bidi/><w:spacing w:after="120"/>')

    def callout(self, text):
        self.p(_runs(text, bold=True, size=26),
               f'<w:pBdr><w:right w:val="single" w:color="{AMBER}" w:sz="18" w:space="8"/></w:pBdr>'
               f'<w:shd w:fill="{CREAM}" w:color="auto" w:val="clear"/><w:bidi/>'
               '<w:spacing w:after="140" w:line="340"/><w:jc w:val="start"/>')

    def h1(self, text, page_break=False):
        pb = "<w:pageBreakBefore/>" if page_break else ""
        self.p(_runs(text, bold=True, color=TEAL, size=36),
               f'<w:pStyle w:val="Heading1"/>{pb}<w:bidi/><w:spacing w:after="160" w:before="320"/>')

    def h2(self, text, optional=False):
        tag = _runs("  للاستزادة", bold=True, color=AMBER, size=20) if optional else ""
        self.p(_runs(text, bold=True, color=TEAL, size=30) + tag,
               '<w:pStyle w:val="Heading2"/><w:keepNext/><w:bidi/><w:spacing w:after="120" w:before="280"/>')

    def para(self, text, size=26, keep=False):
        kn = "<w:keepNext/>" if keep else ""
        self.p(_runs(text, size=size, body=True), f'{kn}<w:bidi/><w:spacing w:after="140" w:line="{LINE}"/><w:jc w:val="start"/>')

    def box(self, lead, text, fill, bar):
        inner = (_runs(lead, bold=True, color=bar, size=26) if lead else "") + _runs(text, size=26, body=True)
        self.p(inner, f'<w:pBdr><w:right w:val="single" w:color="{bar}" w:sz="18" w:space="8"/></w:pBdr>'
                      f'<w:shd w:fill="{fill}" w:color="auto" w:val="clear"/><w:bidi/>'
                      f'<w:spacing w:after="160" w:line="{LINE}"/><w:jc w:val="start"/>')

    def example(self, text):
        self.box("مثال محلول: ", text, EX_FILL, EX_BAR)

    def advanced(self, adv):
        """Doctoral box: debates, readings and an assignment."""
        bdr = (f'<w:pBdr><w:right w:val="single" w:color="{ADV_BAR}" w:sz="18" w:space="8"/></w:pBdr>'
               f'<w:shd w:fill="{ADV_FILL}" w:color="auto" w:val="clear"/><w:bidi/>')
        self.p(_runs("للدراسات المتقدمة (الدكتوراه)", bold=True, color=ADV_BAR, size=28),
               bdr + '<w:keepNext/><w:spacing w:before="200" w:after="60"/><w:jc w:val="start"/>')
        for t in adv.get("paras", []):
            self.p(_runs(t, size=24, body=True), bdr + f'<w:spacing w:after="60" w:line="{LINE}"/><w:jc w:val="start"/>')
        if adv.get("questions"):
            self.p(_runs("أسئلة بحثية مفتوحة", bold=True, color=ADV_BAR, size=24), bdr + '<w:keepNext/><w:spacing w:after="40"/>')
            for q in adv["questions"]:
                self.p(_runs("◂ " + q, size=24, body=True), bdr + f'<w:spacing w:after="40" w:line="{LINE}"/><w:jc w:val="start"/>')
        if adv.get("readings"):
            self.p(_runs("قراءات محكّمة للنقد والعرض", bold=True, color=ADV_BAR, size=24), bdr + '<w:keepNext/><w:spacing w:after="40"/>')
            ltr = bdr.replace("<w:bidi/>", "")
            for r in adv["readings"]:
                self.p(f'<w:r>{_rpr(latin=True, size=20)}<w:t xml:space="preserve">{escape(r)}</w:t></w:r>',
                       ltr + '<w:spacing w:after="40"/><w:ind w:left="360" w:hanging="360"/><w:jc w:val="left"/>')
        if adv.get("task"):
            self.p(_runs("تكليف الدكتوراه: ", bold=True, color=ADV_BAR, size=24) + _runs(adv["task"], size=24, body=True),
                   bdr + f'<w:spacing w:after="200" w:line="{LINE}"/><w:jc w:val="start"/>')

    def label(self, text):
        self.p(_runs(text, bold=True, color=TEAL, size=26),
               '<w:keepNext/><w:bidi/><w:spacing w:after="80" w:before="120"/><w:jc w:val="start"/>')

    def teal_line(self, text):
        self.p(_runs(text, bold=True, color=TEAL, size=26),
               '<w:bidi/><w:spacing w:after="140" w:line="340"/><w:jc w:val="start"/>')

    def lead_para(self, lead, text, lead_color=TEAL):
        inner = _runs(lead, bold=True, color=lead_color, size=26) + _runs(text, size=26, body=True)
        self.p(inner, f'<w:bidi/><w:spacing w:after="140" w:line="{LINE}"/><w:jc w:val="start"/>')

    def items(self, items, numbered=True):
        nid = self.new_num(3) if numbered else 3
        for it in items:
            self.p(_runs(it, size=26, body=True),
                   f'<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{nid}"/></w:numPr>'
                   '<w:bidi/><w:spacing w:after="80" w:line="340"/>')

    def small(self, text, color=GREY, size=22, bold=False):
        self.p(_runs(text, size=size, color=color, bold=bold, body=True),
               '<w:bidi/><w:spacing w:after="120"/><w:jc w:val="start"/>')

    def blank(self):
        self.p("", '<w:jc w:val="start"/>')

    def image(self, path, width_in=2.5, caption=None):
        par = self.doc.add_paragraph()
        par.add_run().add_picture(path, width=Inches(width_in))
        ppr = par._p.get_or_add_pPr()
        ppr.append(parse_xml(f'<w:spacing {W_NS} w:after="60" w:before="120"/>'))
        ppr.append(parse_xml(f'<w:jc {W_NS} w:val="center"/>'))
        # python-docx appends at the end of body before sectPr already
        if caption:
            self.p(_runs(caption, size=21, color=GREY), '<w:bidi/><w:spacing w:after="200"/><w:jc w:val="center"/>')

    def table(self, header, rows, widths=None, caption=None):
        if caption:
            self.tbl_n += 1
            self.p(_runs(f"جدول {self.ch['num']}\u200e-{self.tbl_n}: ", bold=True, color=TEAL, size=21) + _runs(caption, size=21, color=GREY),
                   '<w:pStyle w:val="TableCaption"/><w:keepNext/><w:bidi/><w:spacing w:before="60" w:after="60"/><w:jc w:val="start"/>')
        header = ["المقابل الإنجليزي" if h.strip() == "Term" else h for h in header]
        n = len(header)
        total = 9300
        if not widths:
            widths = [total // n] * n
            if n == 3:
                widths = [2200, 2200, 4900]
        else:
            s = sum(widths)
            widths = [int(w * total / s) for w in widths]
        bord = ''.join(f'<w:{s} w:val="single" w:color="B0B8BA" w:sz="4"/>' for s in ("top", "left", "bottom", "right"))
        mar = ('<w:tcMar><w:top w:type="dxa" w:w="70"/><w:left w:type="dxa" w:w="100"/>'
               '<w:bottom w:type="dxa" w:w="70"/><w:right w:type="dxa" w:w="100"/></w:tcMar>')

        def cell(text, w, head):
            shd = f'<w:shd w:fill="{TEAL}" w:color="auto" w:val="clear"/>' if head else ""
            runs = _runs(str(text), bold=head, color="FFFFFF" if head else None, size=22, explicit_nobold=not head, body=not head)
            return (f'<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="{w}"/><w:tcBorders>{bord}</w:tcBorders>{shd}{mar}</w:tcPr>'
                    f'<w:p><w:pPr><w:bidi/></w:pPr>{runs}</w:p></w:tc>')

        grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
        tb = ''.join(f'<w:{s} w:val="single" w:color="auto" w:sz="4"/>' for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
        xml = [f'<w:tbl {W_NS}><w:tblPr><w:bidiVisual/><w:tblW w:type="dxa" w:w="{total}"/><w:tblBorders>{tb}</w:tblBorders></w:tblPr>'
               f'<w:tblGrid>{grid}</w:tblGrid>']
        xml.append('<w:tr><w:trPr><w:tblHeader/></w:trPr>' + "".join(cell(h, w, True) for h, w in zip(header, widths)) + '</w:tr>')
        for r in rows:
            xml.append('<w:tr><w:trPr><w:cantSplit/></w:trPr>' + "".join(cell(c, w, False) for c, w in zip(r, widths)) + '</w:tr>')
        xml.append("</w:tbl>")
        self._add("".join(xml))
        self.blank()

    def page_break(self):
        self.p('<w:r><w:br w:type="page"/></w:r>', "")

    def save(self, path):
        self.doc.core_properties.title = f"الفصل {self.ch['num']}: {self.ch['title']}"
        self.doc.core_properties.author = "د. عثمان نصر"
        from normalize import normalize_document
        normalize_document(self.doc)
        self.doc.save(path)


# ---------------------------------------------------------------------------
def block_to_doc(b, blk, title=""):
    kind = blk[0]
    if kind == "table":
        _, header, rows = blk[:3]
        widths = blk[3] if len(blk) > 3 else None
        b.table(header, rows, widths, caption=title)
    elif kind == "two":
        for title, items in blk[1:3]:
            b.label(title)
            b.items(items, numbered=False)
    elif kind in ("steps", "grid"):
        b.table(["العنصر", "Term", "الشرح"], [[a, e, d] for a, e, d in blk[1]], caption=title)
    else:
        raise ValueError(kind)


def body_to_doc(b, body, title=""):
    """Render an ordered section body (used by chapters converted from the earlier generator)."""
    first_table = True
    for e in body:
        k = e[0]
        if k == "para":
            b.para(e[1])
        elif k == "table":
            b.table(e[1], e[2], e[3] if len(e) > 3 else None, caption=title if first_table else title + " (تابع)")
            first_table = False
        elif k == "label":
            b.label(e[1])
        elif k == "list":
            b.items(e[1], numbered=(e[2] != "bullet"))
        elif k == "callout":
            b.callout(e[1])
        elif k == "ex":
            b.example(e[1])
        elif k == "teal":
            b.teal_line(e[1])
        elif k == "small":
            b.small(e[1])
        elif k in ("bold", "ref"):
            b.small(e[1], color=None, size=26 if k == "bold" else 21, bold=True)
        elif k == "task":
            b.para(e[1], size=23)
        else:
            raise ValueError(e)


def section_to_doc(b, s):
    if "body" in s:
        body_to_doc(b, s["body"], s["t"])
        return
    for t in s["p"]:
        b.para(t)
    if s.get("blk"):
        block_to_doc(b, s["blk"], s["t"])
    if s.get("take"):
        b.callout(s["take"])
    if s.get("ex"):
        b.example(s["ex"])


def build_doc(ch, out_path, img_dir):
    if "advanced" not in ch:
        try:
            from advanced import ADV
            if ch["num"] in ADV:
                ch = dict(ch, advanced=ADV[ch["num"]])
        except ImportError:
            pass
    ch = dedupe_english(walk_terms(ch))
    b = DocBuilder(ch)
    n = ch["num"]
    b.kicker(f"الفصل {ORD[n]}")
    b.title(ch["title"])
    b.callout(f"سؤال الفصل: {ch['question']}")
    b.image(os.path.join(img_dir, ch["cover_img"]), 2.5, caption=f"الشكل {n}\u200e-1: {ch['title']}")

    oc = ch["opening"]
    b.h1(f"حالة افتتاحية: {oc['title']}")
    for t in oc["paras"]:
        b.para(t)
    b.callout(oc["question"])

    b.h1("مخرجات التعلم")
    b.para("بعد دراسة هذا الفصل يستطيع الطالب أن:", keep=True)
    b.items([f"{ar}: {txt}" for ar, en, txt in ch["outcomes"]])
    if ch.get("advanced", {}).get("outcome"):
        b.lead_para("مخرج إضافي للدكتوراه: ", ch["advanced"]["outcome"], lead_color="6B4E71")
    b.table(["المخرج", "مجال إطار المؤهلات الوطني"],
            [[f"{ar}: {txt}", nqf_domain(ar)] for ar, en, txt in ch["outcomes"]]
            + ([[f"(دكتوراه) {ch['advanced']['outcome']}", nqf_domain(ch['advanced']['outcome'].split()[0])]]
               if ch.get("advanced", {}).get("outcome") else []), [3, 1.3],
            caption="ارتباط مخرجات الفصل بمجالات إطار المؤهلات الوطني")

    b.h1("خريطة الفصل: 4 محاور ومخرج واحد")
    b.table(["المحور", "السؤال الذي يجيب عنه", "الموضوعات"],
            [[f"{ax['letter']}. {ax['name']}", ax["q"], ax["topics"]] for ax in ch["axes"]], caption="خريطة الفصل")
    b.callout(f"المخرج: {ch['output']}")

    current_axis = None
    for i, s in enumerate(ch["sections"], 1):
        if s["axis"] != current_axis:
            current_axis = s["axis"]
            ax = ch["axes"][current_axis]
            b.h1(f"المحور {ax['letter']}: {ax['name']}")
        b.h2(f"{n}.{i} {s['t']}", optional=s.get("opt", False))
        section_to_doc(b, s)

    if ch.get("advanced"):
        b.h1("للدراسات المتقدمة")
        b.advanced(ch["advanced"])

    app = ch["app"]
    b.h1(f"التطبيق: {app['title']}")
    b.items(app["steps"])
    b.small(f"الملف: {app['file']}   |   الزمن المقترح: {app['time']}")

    b.h1("حالة للنقاش ومهمة بحثية")
    b.items(ch["discussion"])
    b.h2("مهمة بحثية قصيرة")
    b.para(ch["task"], size=23)

    b.h1("خلاصة الفصل")
    for (lab, txt) in ch["summary"]:
        b.lead_para(f"{lab}: ", txt)
    b.callout(ch["next"])

    b.h1("أسئلة المراجعة")
    for k, (lvl, q) in enumerate(ch["review"], 1):
        b.teal_line(f"{k}. [{lvl}] {q}")

    b.h1("المصطلحات")
    b.table(["المصطلح", "بالإنجليزية"], [[a, e] for a, e in ch["terms"]], [1, 1])

    b.h1("المراجع")
    for ref, note in ch["refs"]:
        b.small(f"{ref} — {note}", color=None, size=21, bold=True)

    if ch.get("appendix"):
        b.page_break()
        b.h1("ملحق: موضوعات للاستزادة")
        b.small("أقسام فنية اختيارية لمن يريد التعمق في الأساس الرياضي أو الإحصائي لما ورد في متن الفصل، ولا تتوقف عليها بطاقة الفصل ولا تمارينه.")
        for k, s in enumerate(ch["appendix"], 1):
            b.h2(f"م-{k} {s['t']}", optional=True)
            section_to_doc(b, s)

    b.page_break()
    b.h1("ملحق: الإجابات النموذجية")
    for k, a in enumerate(ch["answers"], 1):
        b.small(f"{k}. {a}", color=None, size=26, bold=True)
    b.save(out_path)
