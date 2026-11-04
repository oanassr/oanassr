"""Render a chapter (content dict) to a Word book chapter matching the Ch01–Ch09 design."""
import copy
import os
from xml.sax.saxutils import escape

import docx
from docx.oxml import parse_xml
from docx.shared import Inches

from textutil import split_runs, ORD

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "..", "templates", "Ch09_Book.docx")

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
TEAL, AMBER, CREAM, GREY = "0F4C5C", "E09F3E", "FFF4DD", "555555"
FONT = '<w:rFonts w:ascii="Cairo" w:cs="Cairo" w:eastAsia="Cairo" w:hAnsi="Cairo"/>'


def _rpr(bold=False, color=None, size=26, latin=False, explicit_nobold=False):
    b = "<w:b/><w:bCs/>" if bold else ('<w:b w:val="false"/><w:bCs w:val="false"/>' if explicit_nobold else "")
    c = f'<w:color w:val="{color}"/>' if color else ""
    rtl = '<w:rtl w:val="false"/>' if latin else "<w:rtl/>"
    return f"<w:rPr>{FONT}{b}{c}<w:sz w:val=\"{size}\"/><w:szCs w:val=\"{size}\"/>{rtl}</w:rPr>"


def _runs(text, **kw):
    out = []
    for seg, latin in split_runs(text):
        out.append(f'<w:r>{_rpr(latin=latin, **kw)}<w:t xml:space="preserve">{escape(seg)}</w:t></w:r>')
    return "".join(out)


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
        self.p(_runs(text, bold=True, color=TEAL, size=52), '<w:bidi/><w:spacing w:after="120"/>')

    def callout(self, text):
        self.p(_runs(text, bold=True, size=26),
               f'<w:pBdr><w:right w:val="single" w:color="{AMBER}" w:sz="18" w:space="8"/></w:pBdr>'
               f'<w:shd w:fill="{CREAM}" w:color="auto" w:val="clear"/><w:bidi/>'
               '<w:spacing w:after="140" w:line="340"/><w:jc w:val="start"/>')

    def h1(self, text, page_break=False):
        pb = "<w:pageBreakBefore/>" if page_break else ""
        self.p(_runs(text, bold=True, color=TEAL, size=36),
               f'<w:pStyle w:val="Heading1"/>{pb}<w:bidi/><w:spacing w:after="160" w:before="320"/>')

    def h2(self, text):
        self.p(_runs(text, bold=True, color=TEAL, size=30),
               '<w:pStyle w:val="Heading2"/><w:keepNext/><w:bidi/><w:spacing w:after="120" w:before="280"/>')

    def para(self, text, size=26, keep=False):
        kn = "<w:keepNext/>" if keep else ""
        self.p(_runs(text, size=size), f'{kn}<w:bidi/><w:spacing w:after="140" w:line="340"/><w:jc w:val="start"/>')

    def label(self, text):
        self.p(_runs(text, bold=True, color=TEAL, size=26),
               '<w:keepNext/><w:bidi/><w:spacing w:after="80" w:before="120"/><w:jc w:val="start"/>')

    def teal_line(self, text):
        self.p(_runs(text, bold=True, color=TEAL, size=26),
               '<w:bidi/><w:spacing w:after="140" w:line="340"/><w:jc w:val="start"/>')

    def lead_para(self, lead, text, lead_color=TEAL):
        inner = _runs(lead, bold=True, color=lead_color, size=26) + _runs(text, size=26)
        self.p(inner, '<w:bidi/><w:spacing w:after="140" w:line="340"/><w:jc w:val="start"/>')

    def items(self, items, numbered=True):
        nid = self.new_num(3) if numbered else 3
        for it in items:
            self.p(_runs(it, size=26),
                   f'<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{nid}"/></w:numPr>'
                   '<w:bidi/><w:spacing w:after="80" w:line="340"/>')

    def small(self, text, color=GREY, size=22, bold=False):
        self.p(_runs(text, size=size, color=color, bold=bold),
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

    def table(self, header, rows, widths=None):
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
            runs = _runs(str(text), bold=head, color="FFFFFF" if head else None, size=22, explicit_nobold=not head)
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
def block_to_doc(b, blk):
    kind = blk[0]
    if kind == "table":
        _, header, rows = blk[:3]
        widths = blk[3] if len(blk) > 3 else None
        b.table(header, rows, widths)
    elif kind == "two":
        for title, items in blk[1:3]:
            b.label(title)
            b.items(items, numbered=False)
    elif kind in ("steps", "grid"):
        b.table(["العنصر", "Term", "الشرح"], [[a, e, d] for a, e, d in blk[1]])
    else:
        raise ValueError(kind)


def build_doc(ch, out_path, img_dir):
    b = DocBuilder(ch)
    n = ch["num"]
    b.kicker(f"الفصل {ORD[n]}")
    b.title(ch["title"])
    b.callout(f"سؤال الفصل: {ch['question']}")
    b.image(os.path.join(img_dir, ch["cover_img"]), 2.5, caption=f"الشكل {n}-1")

    oc = ch["opening"]
    b.h1(f"حالة افتتاحية: {oc['title']}")
    for t in oc["paras"]:
        b.para(t)
    b.callout(oc["question"])

    b.h1("مخرجات التعلم")
    b.para("بعد دراسة هذا الفصل يستطيع الطالب أن:", keep=True)
    b.items([f"{ar} ({en}): {txt}" for ar, en, txt in ch["outcomes"]])

    b.h1("خريطة الفصل: 4 محاور ومخرج واحد")
    b.table(["المحور", "السؤال الذي يجيب عنه", "الموضوعات"],
            [[f"{ax['letter']}. {ax['name']}", ax["q"], ax["topics"]] for ax in ch["axes"]])
    b.callout(f"المخرج: {ch['output']}")

    current_axis = None
    for i, s in enumerate(ch["sections"], 1):
        if s["axis"] != current_axis:
            current_axis = s["axis"]
            ax = ch["axes"][current_axis]
            b.h1(f"المحور {ax['letter']}: {ax['name']} ({ax['en']})")
        b.h2(f"{n}.{i} {s['t']}")
        for t in s["p"]:
            b.para(t)
        if s.get("blk"):
            block_to_doc(b, s["blk"])
        if s.get("take"):
            b.callout(s["take"])
        if s.get("ex"):
            b.lead_para("مثال محلول: ", s["ex"])

    app = ch["app"]
    b.h1(f"التطبيق: {app['title']}")
    b.items(app["steps"])
    b.small(f"الملف: {app['file']}   |   الزمن المقترح: {app['time']}")

    b.h1("حالة للنقاش ومهمة بحثية")
    b.items(ch["discussion"])
    b.h2("مهمة بحثية قصيرة (Mini Research Task)")
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

    b.page_break()
    b.h1("ملحق: الإجابات النموذجية")
    for k, a in enumerate(ch["answers"], 1):
        b.small(f"{k}. {a}", color=None, size=26, bold=True)
    b.save(out_path)
