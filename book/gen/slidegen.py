"""Render a chapter (content dict) to a 16:9 slide deck matching the Ch01–Ch09 design.

The Ch09 deck is used only for its master, layouts (DARK cover background, LIGHT content
background with footer) – every slide is rebuilt from the chapter data.
"""
import copy
import math
import os
import re

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from textutil import ORD

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "..", "templates", "Ch09_Slides.pptx")

TEAL, TEAL2, AMBER, RUST, NAVY, INK = "0F4C5C", "2C7A7B", "E09F3E", "B5532A", "3D5A80", "14213D"
PLUM, GOLD, CREAM, SAND, SAND2 = "6B4E71", "B07A1E", "FBF7EF", "F1E9DA", "F6F1E6"
CARD_COLORS = [TEAL, TEAL2, AMBER, RUST, NAVY, PLUM, TEAL2, AMBER]
AXIS_TAB = [TEAL, TEAL2, AMBER, RUST]
AXIS_KICK = [TEAL, TEAL2, GOLD, RUST]


def rgb(h):
    return RGBColor.from_string(h)


def est_lines(text, width_in, size_pt):
    cpl = max(1.0, width_in * 72 / (size_pt * 0.47))
    lines = 0
    for part in str(text).split("\n"):
        lines += max(1, math.ceil(len(part) / cpl))
    return lines


def fit_size(text, w, h, size, min_size=8.0, line=1.32):
    s = size
    while s > min_size and est_lines(text, w, s) * s * line / 72 > h:
        s -= 0.5
    return s


class Deck:
    def __init__(self, ch):
        self.ch = ch
        self.prs = Presentation(TEMPLATE)
        prs = self.prs
        sld_ids = prs.slides._sldIdLst
        for sid in list(sld_ids):
            prs.part.drop_rel(sid.rId)
            sld_ids.remove(sid)
        self.layout_dark = prs.slide_layouts[1]
        self.layout_light = prs.slide_layouts[2]
        footer = f"إدارة اتخاذ القرار على أساس البيانات  –  الفصل {ORD[ch['num']]}: {ch['title']}"
        for t in self.layout_light.element.iter(qn("a:t")):
            if "إدارة اتخاذ القرار" in (t.text or ""):
                t.text = footer
        self.n = 0

    # ---------- primitives ----------
    def box(self, s, x, y, w, h, fill, line=None, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06):
        sh = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
        if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
            sh.adjustments[0] = radius
        sh.fill.solid()
        sh.fill.fore_color.rgb = rgb(fill)
        sh.line.color.rgb = rgb(line or fill)
        sh.line.width = Pt(0.75)
        sh.shadow.inherit = False
        sh.text_frame.text = ""
        return sh

    def text(self, s, x, y, w, h, txt, size, color=INK, bold=False, align="r", anchor="m",
             fit=True, bullets=False, min_size=8.0):
        tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        for side in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
            setattr(tf, side, 0)
        tf.vertical_anchor = {"m": MSO_ANCHOR.MIDDLE, "t": MSO_ANCHOR.TOP}[anchor]
        lines = txt if isinstance(txt, list) else [txt]
        if fit:
            joined = "\n".join(("• " + l) if bullets else l for l in lines)
            size = fit_size(joined, w, h, size, min_size)
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = {"r": PP_ALIGN.RIGHT, "c": PP_ALIGN.CENTER, "l": PP_ALIGN.LEFT}[align]
            pPr = p._p.get_or_add_pPr()
            if align != "l":
                pPr.set("rtl", "1")
            if bullets:
                pPr.set("marL", "152400")
                pPr.set("indent", "-152400")
                spc = etree.SubElement(pPr, qn("a:spcAft"))
                etree.SubElement(spc, qn("a:spcPts")).set("val", "200")
                etree.SubElement(pPr, qn("a:buSzPct")).set("val", "100000")
                etree.SubElement(pPr, qn("a:buChar")).set("char", "•")
            r = p.add_run()
            r.text = str(line)
            f = r.font
            f.size = Pt(size)
            f.bold = bold
            f.color.rgb = rgb(color)
            f.name = "Cairo"
            rPr = r._r.get_or_add_rPr()
            rPr.set("lang", "ar-SA")
            for tag in ("a:ea", "a:cs"):
                el = rPr.find(qn(tag))
                if el is None:
                    el = etree.SubElement(rPr, qn(tag))
                el.set("typeface", "Cairo")
        return tb

    def new(self, dark=False):
        s = self.prs.slides.add_slide(self.layout_dark if dark else self.layout_light)
        for ph in list(s.placeholders):
            ph._element.getparent().remove(ph._element)
        self.n += 1
        if not dark:
            self.text(s, 0.40, 5.28, 0.50, 0.25, str(self.n), 9, color="5B6770", align="l", fit=False)
        return s

    def nav(self, s, active):
        axes = self.ch["axes"]
        for i, x in enumerate([7.35, 5.05, 2.75, 0.45]):
            label = f"{axes[i]['letter']} – {axes[i]['name']}"
            if i == active:
                c = AXIS_TAB[i]
                self.box(s, x, 0.13, 2.20, 0.27, c, radius=0.5)
                self.text(s, x, 0.13, 2.20, 0.27, label, 9.5, color=INK if c == AMBER else "FFFFFF", bold=True, align="c", min_size=7.5)
            else:
                self.box(s, x, 0.13, 2.20, 0.27, CREAM, line="C9C2B3", radius=0.5)
                self.text(s, x, 0.13, 2.20, 0.27, label, 9.5, color="8A8576", align="c", min_size=7.5)

    def header(self, s, kicker, title, en, kick_color=TEAL):
        self.text(s, 0.45, 0.50, 9.10, 0.25, kicker, 10, color=kick_color, bold=True, fit=False)
        self.text(s, 0.45, 0.72, 9.10, 0.50, title, 25, color=INK, bold=True, min_size=17)
        self.text(s, 0.45, 1.20, 9.10, 0.25, en, 12, color=GOLD, min_size=9)

    def lead(self, s, txt, h=0.70):
        self.box(s, 0.45, 1.55, 9.10, h, "FFFFFF")
        self.text(s, 0.65, 1.59, 8.70, h - 0.08, txt, 11.5)

    def bar(self, s, txt, y=4.68, h=0.50, size=10.5, bold=True):
        self.box(s, 0.45, y, 9.10, h, SAND)
        self.text(s, 0.60, y, 8.80, h, txt, size, color=TEAL, bold=bold, align="c")

    def numbered_rows(self, s, items, y0, h, gap, size=12.5, circle=0.52, num_size=16):
        for i, it in enumerate(items):
            y = y0 + i * (h + gap)
            self.box(s, 0.45, y, 9.10, h, "FFFFFF")
            c = [TEAL, TEAL2, AMBER, RUST][i % 4]
            cy = y + (h - circle) / 2
            o = self.box(s, 8.95, cy, circle, circle, c, line="FFFFFF", shape=MSO_SHAPE.OVAL)
            self.text(s, 8.95, cy, circle, circle, str(i + 1), num_size, color=INK if c == AMBER else "FFFFFF", bold=True, align="c", fit=False)
            self.text(s, 0.65, y, 8.10, h, it, size)

    # ---------- content blocks ----------
    def table(self, s, header, rows, widths=None, y=2.37, max_h=2.22):
        n_rows = len(rows) + 1
        cols = len(header)
        rh = min(0.38, max_h / n_rows)
        size = 9.5 if rh >= 0.34 else 8.5
        if not widths:
            widths = [1] * cols
        tot = sum(widths)
        ws = [9.10 * w / tot for w in widths]
        # make sure long cells fit: grow row height estimate
        gf = s.shapes.add_table(n_rows, cols, Inches(0.45), Inches(y), Inches(9.10), Inches(rh * n_rows))
        tbl = gf.table
        tblPr = tbl._tbl.tblPr
        for attr in ("firstRow", "bandRow"):
            tblPr.set(attr, "0")
        style = tblPr.find(qn("a:tableStyleId"))
        if style is not None:
            tblPr.remove(style)
        # RTL: first logical column on the right
        order = list(range(cols))[::-1]
        for j, w in enumerate(ws[::-1]):
            tbl.columns[j].width = Inches(w)
        for r in range(n_rows):
            tbl.rows[r].height = Inches(rh)
            data = header if r == 0 else rows[r - 1]
            for j, src in enumerate(order):
                cell = tbl.cell(r, j)
                fill = TEAL if r == 0 else ("FFFFFF" if r % 2 else SAND2)
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb(fill)
                cell.margin_left = cell.margin_right = Inches(0.08)
                cell.margin_top = cell.margin_bottom = Inches(0.03)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                txt = str(data[src])
                csize = fit_size(txt, ws[::-1][j] - 0.16, rh - 0.04, size, 7)
                tf = cell.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.alignment = PP_ALIGN.RIGHT
                p._p.get_or_add_pPr().set("rtl", "1")
                run = p.add_run()
                run.text = txt
                run.font.size = Pt(csize)
                run.font.bold = r == 0
                run.font.name = "Cairo"
                run.font.color.rgb = rgb("FFFFFF" if r == 0 else INK)
                rPr = run._r.get_or_add_rPr()
                for tag in ("a:ea", "a:cs"):
                    el = etree.SubElement(rPr, qn(tag))
                    el.set("typeface", "Cairo")
                tcPr = cell._tc.get_or_add_tcPr()
                for ln in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
                    e = etree.SubElement(tcPr, qn(ln))
                    e.set("w", "9525")
                    sf = etree.SubElement(e, qn("a:solidFill"))
                    etree.SubElement(sf, qn("a:srgbClr")).set("val", "D6CFC0")
                # lines must precede fill in tcPr
                fill_el = tcPr.find(qn("a:solidFill"))
                if fill_el is not None:
                    tcPr.remove(fill_el)
                    tcPr.append(fill_el)

    def two(self, s, left, right):
        # right card = first group (teal), left card = second group (rust)
        for (title, items), x, c in ((left, 5.05, TEAL), (right, 0.45, RUST)):
            self.box(s, x, 2.35, 4.50, 2.25, "FFFFFF")
            self.box(s, x, 2.35, 4.50, 0.50, c)
            self.text(s, x + 0.15, 2.35, 4.20, 0.50, title, 12.5, color="FFFFFF", bold=True)
            self.text(s, x + 0.15, 2.95, 4.20, 1.55, items, 12, bullets=True, anchor="t")

    def steps(self, s, items, y=2.40, h=2.15):
        n = len(items)
        gap = 0.27
        w = (9.10 - gap * (n - 1)) / n
        for i, (ar, en, d) in enumerate(items):
            x = 9.55 - w - i * (w + gap)
            c = CARD_COLORS[i % len(CARD_COLORS)]
            fg, sub = (INK, "6B4A00") if c == AMBER else ("FFFFFF", "F4D9A8")
            self.box(s, x, y, w, h, c)
            self.text(s, x + w - 0.50, y + 0.06, 0.40, 0.30, str(i + 1), 13, color=sub, bold=True, align="c", fit=False)
            self.text(s, x + 0.10, y + 0.36, w - 0.20, 0.50, ar, 12, color=fg, bold=True)
            self.text(s, x + 0.10, y + 0.86, w - 0.20, 0.20, en, 8.5, color=sub)
            self.text(s, x + 0.10, y + 1.10, w - 0.20, h - 1.16, d, 10.5, color=fg, anchor="t")
            if i < n - 1:
                tri = s.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(x - 0.235), Inches(y + h / 2 - 0.12), Inches(0.20), Inches(0.24))
                tri.rotation = 270
                tri.fill.solid()
                tri.fill.fore_color.rgb = rgb("5B6770")
                tri.line.color.rgb = rgb("5B6770")

    def grid(self, s, items, y=2.27, bottom=5.12):
        n = len(items)
        cols = 3 if n in (3, 5, 6, 9) else 4 if n in (4, 7, 8) else 3
        if n == 4:
            cols = 2
        rows = math.ceil(n / cols)
        gapx, gapy = 0.10, 0.10
        w = (9.10 - gapx * (cols - 1)) / cols
        h = min(1.40, (bottom - y - gapy * (rows - 1)) / rows)
        for i, (ar, en, d) in enumerate(items):
            r, c_ = divmod(i, cols)
            x = 9.55 - w - c_ * (w + gapx)
            yy = y + r * (h + gapy)
            c = CARD_COLORS[i % len(CARD_COLORS)]
            fg, sub = (INK, "6B4A00") if c == AMBER else ("FFFFFF", "F4D9A8")
            self.box(s, x, yy, w, h, c)
            self.text(s, x + w - 0.50, yy + 0.06, 0.40, 0.32, str(i + 1), 14, color=sub, bold=True, align="c", fit=False)
            self.text(s, x + 0.10, yy + 0.04, w - 0.62, 0.36, ar, 12.5, color=fg, bold=True)
            self.text(s, x + 0.10, yy + 0.41, w - 0.20, 0.20, en, 8.5, color=sub)
            self.text(s, x + 0.10, yy + 0.64, w - 0.20, h - 0.70, d, 10.5, color=fg, anchor="t")

    # ---------- slide types ----------
    def cover(self, img):
        ch = self.ch
        s = self.new(dark=True)
        self.text(s, 3.20, 0.90, 6.30, 0.40, f"الفصل {ORD[ch['num']]}", 16, color=AMBER, bold=True, fit=False)
        self.text(s, 3.20, 1.35, 6.30, 1.00, ch["title"], 40, color="FFFFFF", bold=True, min_size=26)
        self.text(s, 3.20, 2.35, 6.30, 0.45, ch["title_en"], 22, color="F4D9A8", min_size=14)
        self.text(s, 3.20, 3.10, 6.30, 0.80, ch["question"], 16, color="DCEBEA", min_size=11)
        self.text(s, 3.20, 4.55, 6.30, 0.30, "إدارة اتخاذ القرار على أساس البيانات  –  MIS6004", 11, color="9FC3C3", fit=False)
        self.text(s, 3.20, 4.85, 6.30, 0.25, "Data-Driven Decision Making  –  Master's Programs  –  King Khalid University", 10, color="9FC3C3", fit=False)
        self.box(s, 0.50, 1.00, 2.60, 2.60, CREAM, line=AMBER, shape=MSO_SHAPE.OVAL)
        s.shapes.add_picture(img, Inches(0.80), Inches(1.55), Inches(2.00), Inches(1.50))
        s.notes_slide.notes_text_frame.text = ch["step_note"]

    def objectives(self):
        ch = self.ch
        s = self.new()
        self.nav(s, 0)
        self.header(s, "أهداف التعلم", "أهداف الفصل", "Learning Objectives")
        self.numbered_rows(s, [f"{ar}: {t}" for ar, en, t in ch["outcomes"]], 1.60, 0.75, 0.10)

    def opening(self, img):
        oc = self.ch["opening"]
        s = self.new()
        self.nav(s, 0)
        self.header(s, "حالة افتتاحية", oc["title"], oc["title_en"])
        s.shapes.add_picture(img, Inches(0.50), Inches(1.70), Inches(3.10), Inches(2.00))
        self.box(s, 3.85, 1.55, 5.70, 3.60, "FFFFFF")
        self.text(s, 4.05, 1.65, 5.30, 1.20, oc["slide"][0], 11, anchor="t")
        self.text(s, 4.05, 2.90, 5.30, 1.05, oc["slide"][1], 11.5, anchor="t")
        self.text(s, 4.05, 4.00, 5.30, 0.62, oc["slide_q"], 12.5, color=RUST, bold=True)

    def section(self, sec):
        s = self.new()
        ax = sec["axis"]
        axd = self.ch["axes"][ax]
        self.nav(s, ax)
        self.header(s, f"{axd['letter']} – {axd['name']}", sec["t"], sec["en"], AXIS_KICK[ax])
        self.lead(s, sec["lead"])
        blk = sec.get("blk")
        kind = blk[0] if blk else None
        if kind == "table":
            widths = blk[3] if len(blk) > 3 else None
            self.table(s, blk[1], blk[2], widths, max_h=2.22 if sec.get("take") else 2.70)
        elif kind == "two":
            self.two(s, blk[1], blk[2])
        elif kind == "steps":
            if len(blk[1]) <= 5:
                self.steps(s, blk[1])
            else:
                self.grid(s, blk[1])
        elif kind == "grid":
            self.grid(s, blk[1], bottom=4.58 if sec.get("take") else 5.12)
        if sec.get("take"):
            self.bar(s, sec["take"])

    def application(self):
        app = self.ch["app"]
        s = self.new()
        self.nav(s, 3)
        self.header(s, "التطبيق", app["title"], app["en"], RUST)
        n = len(app["steps"])
        h = min(0.60, (3.25 - 0.06 * (n - 1)) / n)
        self.numbered_rows(s, app["steps"], 1.55, h, 0.06, size=11.5)
        self.bar(s, f"الملف: {app['file']}   |   الزمن المقترح: {app['time']}", y=4.85, h=0.38, size=9.5, bold=False)

    def summary(self):
        ch = self.ch
        s = self.new()
        self.nav(s, 3)
        self.header(s, "خلاصة", "خلاصة الفصل", ch["title_en"], RUST)
        self.numbered_rows(s, [t for _, t in ch["summary"]], 1.55, 0.72, 0.08, size=11.5, circle=0.48, num_size=14)
        self.bar(s, ch["next"], y=4.78, h=0.42)

    def references(self):
        refs = self.ch["refs"]
        s = self.new()
        self.nav(s, 3)
        self.header(s, "المصادر", "المراجع", "References", RUST)
        n = len(refs)
        h = min(0.29, (3.65 - 0.03 * (n - 1)) / n)
        for i, (ref, note) in enumerate(refs):
            y = 1.50 + i * (h + 0.03)
            self.box(s, 0.45, y, 9.10, h, "FFFFFF" if i % 2 == 0 else SAND2, radius=0.15)
            self.text(s, 0.55, y, 6.50, h, ref, 8.5, align="l", min_size=6.5)
            self.text(s, 7.15, y, 2.30, h, note, 9, min_size=7)

    def save(self, path):
        self.prs.core_properties.title = f"الفصل {self.ch['num']}: {self.ch['title']}"
        self.prs.save(path)


def build_deck(ch, out_path, img_dir):
    d = Deck(ch)
    d.cover(os.path.join(img_dir, ch["cover_img"]))
    d.objectives()
    d.opening(os.path.join(img_dir, ch["case_img"]))
    for sec in ch["sections"]:
        d.section(sec)
    d.application()
    d.summary()
    d.references()
    d.save(out_path)
