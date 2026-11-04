"""Build chapter files: python3 build.py 10 11 12 13  (outputs to ../out/)."""
import importlib
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(HERE, "content"))

from docgen import build_doc  # noqa: E402
from slidegen import build_deck  # noqa: E402

OUT = os.path.join(HERE, "..", "out")
IMG = os.path.join(HERE, "..", "img")


def build(num):
    mod = importlib.import_module(f"ch{num:02d}")
    ch = mod.CH
    os.makedirs(os.path.join(OUT, "chapter"), exist_ok=True)
    os.makedirs(os.path.join(OUT, "slides"), exist_ok=True)
    os.makedirs(os.path.join(OUT, "excel"), exist_ok=True)
    build_doc(ch, os.path.join(OUT, "chapter", f"Ch{num:02d}_Book.docx"), IMG)
    build_deck(ch, os.path.join(OUT, "slides", f"Ch{num:02d}_Slides.pptx"), IMG)
    mod.build_xlsx(os.path.join(OUT, "excel", ch["app"]["file"]))
    print("built chapter", num, "-", len(ch["sections"]), "sections")


if __name__ == "__main__":
    for a in sys.argv[1:] or ["10"]:
        build(int(a))
