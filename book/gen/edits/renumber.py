"""Renumber chapters 5-13 to the 18-chapter order and rewrite cross-references.

old -> new: 1-4 same, 5->6, 6->7, 7->8, 8->9, 9->10, 10->11, 11->13, 12->14, 13->18
Run once from book/gen. Content files are renamed (git mv) and every "الفصل X" reference
(ordinal words or digits, including "الفصلين X وY" and ranges "الفصول a–b") is mapped.
"""
import os, re, subprocess, sys
GEN = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, GEN)
from textutil import ORD

MAP = {1: 1, 2: 2, 3: 3, 4: 4, 5: 6, 6: 7, 7: 8, 8: 9, 9: 10, 10: 11, 11: 13, 12: 14, 13: 18}
W2N = {v: k for k, v in ORD.items() if k <= 13}
ORDS = sorted(W2N, key=len, reverse=True)
ORD_RE = "(?:" + "|".join(map(re.escape, ORDS)) + ")"
TOKEN = rf"(?:{ORD_RE}|\d+)"
REF = re.compile(rf"((?:[وفبل]|لل)?(?:الفصل|الفصلين|الفصول|فصل|فصلي))(\s+)({TOKEN}(?:\s*(?:و|إلى|،|–|-)\s*(?:ال)?{TOKEN})*)")
ONE = re.compile(rf"{ORD_RE}|\d+")


def map_tokens(seq):
    def r(m):
        t = m.group(0)
        if t.isdigit():
            n = int(t)
            return str(MAP.get(n, n))
        return ORD[MAP[W2N[t]]]
    return ONE.sub(r, seq)


def remap(text):
    return REF.sub(lambda m: m.group(1) + m.group(2) + map_tokens(m.group(3)), text)


if __name__ == "__main__":
    cdir = os.path.join(GEN, "content")
    # read all, rewrite, then rename from the top down so names never collide
    for old in sorted(MAP, reverse=True):
        new = MAP[old]
        src_p = os.path.join(cdir, f"ch{old:02d}.py")
        s = open(src_p, encoding="utf-8").read()
        if "'source':" in s:  # converted chapters: rewrite the dict (strings are wrapped across lines)
            import importlib, pprint
            sys.path.insert(0, cdir)
            from textterms import walk
            ch = importlib.import_module(f"ch{old:02d}").CH
            ch = walk(ch, remap)
            ch["num"] = new
            s = s.partition("CH = ")[0] + "CH = " + pprint.pformat(ch, width=150, sort_dicts=False) + "\n\nbuild_xlsx = None\n"
        else:
            s = remap(s)
            s = re.sub(r"(?m)^(    num=)%d\b" % old, lambda m: m.group(1) + str(new), s, count=1)
        s = s.replace(f"ch{old:02d}_cover.png", f"ch{new:02d}_cover.png").replace(f"ch{old:02d}_case.png", f"ch{new:02d}_case.png")
        open(src_p, "w", encoding="utf-8").write(s)
        if new != old:
            subprocess.run(["git", "mv", src_p, os.path.join(cdir, f"ch{new:02d}.py")], check=True)
            for kind in ("cover", "case"):
                ip = os.path.join(GEN, "..", "img", f"ch{old:02d}_{kind}.png")
                if os.path.exists(ip):
                    subprocess.run(["git", "mv", ip, os.path.join(GEN, "..", "img", f"ch{new:02d}_{kind}.png")], check=True)
        print(old, "->", new)
