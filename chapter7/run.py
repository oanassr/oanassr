import sqlite3, re, sys
db = sqlite3.connect(":memory:"); db.executescript(open("schema.sql").read())
out = [f"# Verified outputs (SQLite {sqlite3.sqlite_version})\n"]
for block in re.split(r"(?m)^-- ", open("queries.sql").read())[1:]:
    title, sql = block.split("\n", 1)
    cur = db.execute(sql.strip().rstrip(";"))
    rows = cur.fetchall(); cols = [c[0] for c in cur.description]
    out.append(f"## {title}\n```sql\n{sql.strip()}\n```\n| " + " | ".join(cols) + " |\n|" + "---|"*len(cols))
    out += ["| " + " | ".join("NULL" if v is None else str(v) for v in r) + " |" for r in rows]
    out.append("")
open("outputs.md", "w").write("\n".join(out)); print("ok", len(out))
