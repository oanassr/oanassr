"""Unified terminology (applied to content files and to the 1-9 decks)."""
import re

RULES = [
    ("صاحب القرار (Decision Owner)، أي متخذ القرار الذي يملك صلاحية اعتماده", "متخذ القرار (Decision Owner)، أي من يملك صلاحية اعتماده"),
    ("أصحاب القرار", "متخذي القرار"),
    ("صاحب القرار", "متخذ القرار"),
    ("صانع القرار", "متخذ القرار"),
    ("انحياز", "تحيّز"),
    ("تحيز", "تحيّز"),
    ("الدقة الإيجابية", "الضبط"),
    ("القيمة p", "القيمة الاحتمالية"),
    ("قيمة p", "قيمة احتمالية"),
    ("تأطير القرار", "صياغة القرار"),
]


def unify(s):
    for a, b in RULES:
        s = s.replace(a, b)
    s = s.replace("تحيّّز", "تحيّز")
    return s


def walk(o, f=unify):
    if isinstance(o, str):
        return f(o)
    if isinstance(o, dict):
        return {k: walk(v, f) for k, v in o.items()}
    if isinstance(o, list):
        return [walk(v, f) for v in o]
    if isinstance(o, tuple):
        return tuple(walk(v, f) for v in o)
    return o


ENG_PAREN = re.compile(r" \(([A-Za-z][A-Za-z\-’'&/ ]*[A-Za-z])\)")


def dedupe_english(ch, skip=("terms", "refs", "title_en", "en", "source", "app", "cover_img", "case_img")):
    """Keep an English gloss in parentheses only at its first occurrence in the chapter."""
    seen = set()

    def f(s):
        def rep(m):
            k = m.group(1).lower()
            if "et al" in k:
                return m.group(0)
            if k in seen:
                return ""
            seen.add(k)
            return m.group(0)
        return ENG_PAREN.sub(rep, s)

    def w(o, key=None):
        if key in skip:
            return o
        if isinstance(o, str):
            return f(o)
        if isinstance(o, dict):
            return {k: w(v, k) for k, v in o.items()}
        if isinstance(o, list):
            return [w(v) for v in o]
        if isinstance(o, tuple):
            return tuple(w(v) for v in o)
        return o
    return w(ch)
