"""Excel import (approval A-7): read ORIGINAL INPUT CELLS ONLY from a v1.1-format workbook.

* Calculated columns / cells are ignored (cached results are never treated as authoritative).
* Input-cell formulas (e.g. ``=DB_N_Encl/2``) are preserved verbatim and evaluated by the engine.
* Missing sheets/tables/columns and unknown columns are reported; no defaults are substituted.
"""
from __future__ import annotations

from openpyxl import load_workbook
from openpyxl.utils import range_boundaries

from calculations import reference_data as R

TABLES = {"loads": ("Load Schedule", "tblLoad"), "transformers": ("Transformer Sizing", "tblTX"),
          "feeders": ("LV Distribution", "tblFdr"), "library": ("Equipment Library", "tblLib"),
          "categories": ("Equipment Library", "tblCat"), "profiles": ("Scenario Analysis", "tblProf"),
          "scenarios": ("Scenario Analysis", "tblScen")}
TABLE_KEY = {"loads": "tblLoad", "transformers": "tblTX", "feeders": "tblFdr", "library": "tblLib", "categories": "tblCat",
             "profiles": "tblProf", "scenarios": "tblScen"}


def _find_table(wb, name):
    for ws in wb.worksheets:
        if name in ws.tables:
            return ws, ws.tables[name]
    return None, None


def import_workbook(path_or_file, keep_blank_rows=False):
    wb = load_workbook(path_or_file, data_only=False)
    report = dict(errors=[], warnings=[], info=[])
    proj = {}
    for key, (sheet, tname) in TABLES.items():
        ws, t = _find_table(wb, tname)
        if t is None:
            report["errors"].append(f"Table {tname} not found - '{key}' not imported")
            proj[key] = []
            continue
        expected_inputs = R.TABLE_COLUMNS[TABLE_KEY[key]]["inputs"]
        expected_calc = set(R.TABLE_COLUMNS[TABLE_KEY[key]]["calculated"])
        c1, r1, c2, r2 = range_boundaries(t.ref)
        heads = [ws.cell(row=r1, column=c).value for c in range(c1, c2 + 1)]
        calc = {tc.name for tc in t.tableColumns if tc.calculatedColumnFormula is not None}
        missing = [h for h in expected_inputs if h not in heads]
        unknown = [h for h in heads if h not in expected_inputs and h not in expected_calc]
        if missing:
            report["errors"].append(f"{tname}: input columns missing {missing} - left blank, NOT defaulted")
        if unknown:
            report["warnings"].append(f"{tname}: unknown columns {unknown} preserved under '_extra'")
        rows = []
        for r in range(r1 + 1, r2 + 1):
            row, extra = {}, {}
            for j, h in enumerate(heads):
                v = ws.cell(row=r, column=c1 + j).value
                if h in calc or h in expected_calc:
                    continue
                if h in expected_inputs:
                    row[h] = v
                else:
                    extra[h] = v
            if extra:
                row["_extra"] = extra
            if keep_blank_rows or any(v not in (None, "") for k, v in row.items() if k != "_extra"):
                rows.append(row)
        proj[key] = rows
        report["info"].append(f"{tname}: {len(rows)} rows imported")
    # design basis (column C, by the Name in column G); derived cells are skipped
    db = {}
    if "Design Basis" in wb.sheetnames:
        ws = wb["Design Basis"]
        names = {}
        for r in range(1, ws.max_row + 1):
            n = ws.cell(row=r, column=7).value
            if isinstance(n, str) and n.startswith("DB_"):
                names[n[3:]] = r
        for it in R.DB_ITEMS:
            if it["derived"]:
                continue
            r = names.get(it["key"])
            if r is None:
                report["errors"].append(f"Design Basis input DB_{it['key']} not found - left blank")
                db[it["key"]] = None
            else:
                db[it["key"]] = ws.cell(row=r, column=3).value
    else:
        report["errors"].append("Design Basis sheet not found")
    proj["design_basis"] = db
    # rating series & summary settings
    series = {}
    if "Lists" in wb.sheetnames:
        ls = wb["Lists"]
        for c in range(1, ls.max_column + 1):
            h = ls.cell(row=4, column=c).value
            key = {"IEC R10 (kVA)": "IEC R10", "IEEE 3-ph (kVA)": "IEEE 3-phase", "User list (kVA)": "User list"}.get(h)
            if key:
                series[key] = [ls.cell(row=r, column=c).value for r in range(5, 45)]
    proj["rating_series"] = series or R.RATING_SERIES
    if not series:
        report["warnings"].append("Rating series not found - v1.1 template series used")
    settings = {}
    if "Summary" in wb.sheetnames:
        ws = wb["Summary"]
        for r in range(1, ws.max_row + 1):
            a = ws.cell(row=r, column=1).value
            if a == "Days per year for annualisation (input)":
                settings["days_per_year"] = ws.cell(row=r, column=2).value
            if a == "Load category breakdown":
                settings["breakdown_scenario"] = ws.cell(row=r, column=4).value
    if "Transformer Sizing" in wb.sheetnames:
        ws = wb["Transformer Sizing"]
        for r in range(1, ws.max_row + 1):
            if ws.cell(row=r, column=1).value == "Centralized comparison: transformer type":
                settings["central_type"] = ws.cell(row=r, column=3).value
    for k in ("days_per_year", "breakdown_scenario", "central_type"):
        if k not in settings:
            report["errors"].append(f"Setting '{k}' not found - left unset")
    proj["settings"] = settings
    exprs = sum(1 for k in ("loads", "transformers", "feeders", "library") for row in proj[k]
                for v in row.values() if isinstance(v, str) and v.startswith("="))
    exprs += sum(1 for v in db.values() if isinstance(v, str) and v.startswith("="))
    if exprs:
        report["info"].append(f"{exprs} input-cell formula(s) preserved for evaluation by the engine")
    return proj, report
