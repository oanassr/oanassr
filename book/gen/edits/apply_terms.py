"""Apply textterms.RULES to all content files (dict rewrite for 1-9, source rewrite for 10+)."""
import importlib, os, pprint, sys
GEN = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, GEN); sys.path.insert(0, os.path.join(GEN, "content"))
from textterms import walk, unify, RULES

for n in range(1, 30):
    p = os.path.join(GEN, "content", f"ch{n:02d}.py")
    if not os.path.exists(p):
        continue
    src = open(p, encoding="utf-8").read()
    mod = importlib.import_module(f"ch{n:02d}")
    if mod.CH.get("source"):
        new = src.partition("CH = ")[0] + "CH = " + pprint.pformat(walk(mod.CH), width=150, sort_dicts=False) + "\n\nbuild_xlsx = None\n"
    else:
        new = unify(src)
    if new != src:
        open(p, "w", encoding="utf-8").write(new)
        print("updated", n)
