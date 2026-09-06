"""
Assembles the deployable index.html from source/template.html + source/data.json
+ the embedded QubicaAMF badge.

Run after extract.py (or after editing template.html directly):
    python3 build.py

Writes ../index.html (repo root), which is what GitHub Pages serves.
"""
import json, os

BASE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(BASE, "template.html")
DATA = os.path.join(BASE, "data.json")
BADGE_B64 = os.path.join(BASE, "assets", "BADGE.b64")
OUT = os.path.join(BASE, "..", "index.html")

with open(TEMPLATE, encoding="utf-8") as f:
    tpl = f.read()

with open(DATA, encoding="utf-8") as f:
    data = json.load(f)
data_min = json.dumps(data, separators=(",", ":"), ensure_ascii=False)

with open(BADGE_B64, encoding="utf-8") as f:
    badge_b64 = f.read().strip()

out = tpl.replace("__DATA_JSON__", data_min)
out = out.replace("__BADGE_B64__", badge_b64)

with open(OUT, "w", encoding="utf-8") as f:
    f.write(out)

print(f"Wrote {os.path.abspath(OUT)} ({len(out):,} bytes)")
