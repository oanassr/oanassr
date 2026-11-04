"""Helpers shared by the generators: splitting mixed Arabic/Latin text into runs."""
import re

# A Latin segment: starts with an optional opening bracket, a Latin letter or a digit,
# and runs through Latin letters, digits and the punctuation that normally sits inside
# English phrases, citations and numbers.
_LATIN = re.compile(r"\(?[A-Za-z0-9][A-Za-z0-9 .,&'/:+\-%=×_()\[\]]*[A-Za-z0-9)%\]]|\(?[A-Za-z0-9]\)?")


def split_runs(text):
    """Return [(segment, is_latin)] covering the whole text in order."""
    out, pos = [], 0
    for m in _LATIN.finditer(text):
        seg = m.group(0)
        # keep an unmatched trailing ')' (opened in Arabic text) out of the Latin run
        if seg.endswith(")") and seg.count("(") < seg.count(")"):
            seg = seg[:-1]
        s, e = m.start(), m.start() + len(seg)
        if s > pos:
            out.append((text[pos:s], False))
        out.append((seg, True))
        pos = e
    if pos < len(text):
        out.append((text[pos:], False))
    return [(t, l) for t, l in out if t]


ORD = {1: "الأول", 2: "الثاني", 3: "الثالث", 4: "الرابع", 5: "الخامس", 6: "السادس",
       7: "السابع", 8: "الثامن", 9: "التاسع", 10: "العاشر", 11: "الحادي عشر",
       12: "الثاني عشر", 13: "الثالث عشر"}
