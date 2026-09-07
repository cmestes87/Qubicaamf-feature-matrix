"""
Rebuilds source/data.json from the two source spreadsheets.

Run this after replacing either spreadsheet in source/spreadsheets/:
    cd source
    python3 extract.py

Then rebuild the deployable site with:
    python3 build.py
"""
import openpyxl, json, os

BASE = os.path.dirname(os.path.abspath(__file__))
NEOVERSE_XLSX = os.path.join(BASE, "spreadsheets", "Neoverse_for_Claude.xlsx")
BESNV_XLSX = os.path.join(BASE, "spreadsheets", "BES_NV_Environments_and_Features_for_Claude.xlsx")
OUT = os.path.join(BASE, "data.json")


def clean(v):
    if v is None:
        return None
    if isinstance(v, str):
        return v.strip()
    return v


# ---- Neoverse ----
# Sheet1 (2), columns: Name | Neoverse Interactive Edition | Neoverse (Stand Alone) | Requires Subscription | Requires BES NV
# Category header rows have a name in col A and every other column empty.
wb = openpyxl.load_workbook(NEOVERSE_XLSX, data_only=True)
ws = wb.worksheets[0]
rows = list(ws.iter_rows(values_only=True))

neoverse_categories = []
cur = None
for r in rows:
    name, interactive, standalone, sub, reqbes = [clean(x) for x in r]
    if name is None and interactive is None:
        continue  # blank separator row
    if interactive is None and standalone is None and sub is None and reqbes is None:
        cur = {"name": name, "items": []}
        neoverse_categories.append(cur)
        continue
    if cur is None:
        continue
    cur["items"].append({
        "name": name,
        "interactive": interactive == "Included",
        "standalone": standalone == "Included",
        "sub": sub == "Yes",
        "neofi": reqbes == "Yes",
    })

# ---- BES NV ----
# Sheet2, columns: Environments | BES NV | Requires Subscription | BES NV Value | Requires Subscription | Neo-Fi | Requires SuperTouch
# Category header rows (e.g. "HyperBowling", "Classic Games") have a name in col A and every other column empty.
# Rows before the first named category header belong to "Environments".
wb2 = openpyxl.load_workbook(BESNV_XLSX, data_only=True)
ws2 = wb2["Sheet2"]
rows2 = list(ws2.iter_rows(values_only=True))

besnv_categories = []
cur = None
for r in rows2:
    name, besnv, besnv_sub, value, value_sub, neofi, supertouch = [clean(x) for x in r]
    if name is None:
        continue
    if name == "Environments" and besnv == "BES NV":
        continue  # header row
    if besnv is None and value is None and besnv_sub is None and value_sub is None and neofi is None and supertouch is None:
        cur = {"name": name, "items": []}
        besnv_categories.append(cur)
        continue
    if cur is None:
        cur = {"name": "Environments", "items": []}
        besnv_categories.append(cur)
    cur["items"].append({
        "name": name,
        "besnv": bool(besnv),
        "besnvSub": bool(besnv_sub),
        "value": bool(value),
        "valueSub": bool(value_sub),
        "neofi": bool(neofi),
        "superTouch": bool(supertouch),
    })

out = {"neoverse": neoverse_categories, "besnv": besnv_categories}
with open(OUT, "w", encoding="utf-8") as f:
    json.dump(out, f, indent=2, ensure_ascii=False)

print("BES NV categories:", [(c["name"], len(c["items"])) for c in besnv_categories])
print("Neoverse categories:", [(c["name"], len(c["items"])) for c in neoverse_categories])
print(f"Wrote {OUT}")
