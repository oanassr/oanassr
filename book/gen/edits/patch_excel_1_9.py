"""Rename and patch the original Excel files of chapters 1-9 for the 18-chapter order.

Edits text inside the xlsx XML directly (charts and formatting untouched):
chapter numbers/ordinals, file names, unified terminology.
Input ../src_excel/ChNN_*.xlsx  ->  output ../out/excel/ChMM_*.xlsx
"""
import glob, os, re, sys, zipfile
GEN = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, GEN); sys.path.insert(0, os.path.join(GEN, "edits"))
from renumber import remap, MAP
from textterms import unify

SRC = os.path.join(GEN, "..", "src_excel")
DST = os.path.join(GEN, "..", "out", "excel")
os.makedirs(DST, exist_ok=True)


def fix_text(s, old, new):
    s = remap(s)
    s = re.sub(r"Ch0%d_" % old, "Ch%02d_" % new, s)
    return unify(s)


for path in sorted(glob.glob(os.path.join(SRC, "Ch0*_*.xlsx"))):
    base = os.path.basename(path)
    old = int(base[2:4]); new = MAP[old]
    out = os.path.join(DST, "Ch%02d_%s" % (new, base[5:]))
    zin = zipfile.ZipFile(path)
    zout = zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename.endswith(".xml") and ("sharedStrings" in item.filename or "worksheets/" in item.filename
                                                or "docProps" in item.filename or "charts/" in item.filename):
            data = fix_text(data.decode("utf-8"), old, new).encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    print(base, "->", os.path.basename(out))
