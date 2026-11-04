"""Recalculate a workbook with LibreOffice and print every formula cell's value (flags errors)."""
import os, subprocess, sys, tempfile
from openpyxl import load_workbook
path = sys.argv[1]
out = tempfile.mkdtemp(dir="/tmp/claude-0")
subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", out, path], capture_output=True, timeout=180)
calc = load_workbook(os.path.join(out, os.path.basename(path)), data_only=True)
src = load_workbook(path)
bad = 0
for ws in src.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                v = calc[ws.title][c.coordinate].value
                if (v is None and "\"\"" not in c.value) or (isinstance(v, str) and v.startswith(("#", "Err"))):
                    bad += 1
                    print("ERR", ws.title, c.coordinate, c.value[:60], v)
                elif "-v" in sys.argv:
                    print(ws.title, c.coordinate, round(v, 4) if isinstance(v, float) else v)
print("formula errors:", bad)
