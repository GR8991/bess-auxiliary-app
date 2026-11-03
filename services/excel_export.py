"""Excel exports.

1. results_workbook(): engineering input + result workbook (static values from the Python engine).
2. v11_workbook(): the project's inputs written into a v1.1-format FORMULA workbook (same generator that
   produced the delivered template) so the calculation can be re-checked independently in Excel.
"""
from __future__ import annotations

import io
import os
import sys
import tempfile

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

from calculations.load_calculations import CALC_COLUMNS
from calculations.transformer_sizing import TX_CALC_COLUMNS
from calculations.lv_distribution import FDR_CALC_COLUMNS
from data_models.project_schema import TABLE_INPUTS, PROJECT_INFO_FIELDS
from services.json_export import clean
from services import engineering_summary as ES

HDR = PatternFill("solid", fgColor="1F3864")
CAL = PatternFill("solid", fgColor="EDEDED")
INP = PatternFill("solid", fgColor="FFF2CC")


def _sheet(wb, title, header, rows, calc_from=None):
    ws = wb.create_sheet(title[:31])
    ws.append(header)
    for j, _ in enumerate(header, 1):
        c = ws.cell(row=1, column=j)
        c.fill, c.font = HDR, Font(color="FFFFFF", bold=True)
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[c.column_letter].width = 14
    for r in rows:
        ws.append([_cell(v) for v in r])
    if calc_from is not None:
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.fill = CAL if c.column > calc_from else INP
    ws.freeze_panes = "B2"
    return ws


def _cell(v):
    v = clean(v)
    if isinstance(v, (dict, list)):
        return str(v)
    return v


def results_workbook(project, res) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Read Me"
    summ = ES.build(project, res)
    lines = [
        "BESS Auxiliary Engineering - input and result export",
        f"Project: {summ['project']['name']}   Revision: {summ['project']['revision']}   Project ID: {summ['project']['project_id']}",
        f"Calculation engine {res['engine_version']} reproducing workbook v{res['source_workbook_version']} methodology.",
        "Values are STATIC results of the Python engine (not live formulas). Yellow = input, grey = calculated.",
        f"Overall disposition: {summ['disposition']}",
        summ["disposition_note"],
        "Screening-level results only; see the 'Unresolved' and 'Limitations' sheets.",
    ]
    for i, l in enumerate(lines, 1):
        ws.cell(row=i, column=1, value=l).font = Font(bold=(i == 1), size=14 if i == 1 else 11)
    ws.column_dimensions["A"].width = 120
    p = summ["project"]
    _sheet(wb, "Project", ["Field", "Value"], [[k, v] for k, v in p.items()])
    db = res["workbook"]["design_basis"]
    chk = res["workbook"]["design_basis_checks"]
    from calculations import reference_data as R
    _sheet(wb, "Design Basis", ["Name", "Parameter", "Value (effective)", "Unit", "Input check"],
           [[f"DB_{i['key']}", i["label"], db.get(i["key"]), i["unit"], chk.get("DB_" + i["key"])] for i in R.DB_ITEMS])
    lib_cols = TABLE_INPUTS["library"]
    _sheet(wb, "Equipment Library", lib_cols + ["Status"],
           [[l.get(c) for c in lib_cols] + [s] for l, s in zip([x for x in project.get("library") or []], res["workbook"]["library_status"])],
           calc_from=len(lib_cols))
    for title, key, cols, wkey, ccols in (("Load Schedule", "loads", TABLE_INPUTS["loads"], "loads", CALC_COLUMNS),
                                          ("Transformers", "transformers", TABLE_INPUTS["transformers"], "transformers", TX_CALC_COLUMNS),
                                          ("LV Feeders", "feeders", TABLE_INPUTS["feeders"], "feeders", FDR_CALC_COLUMNS)):
        rows = []
        for raw, out in zip(project.get(key) or [], res["workbook"][wkey]):   # raw inputs keep '=...' expressions
            if raw.get("_padding"):
                continue
            rows.append([raw.get(c) for c in cols] + [out.get(c) for c in ccols])
        _sheet(wb, title, cols + list(ccols), rows, calc_from=len(cols))
    sc = res["workbook"]["scenarios"]["rows"]
    keys = ["Code", "Scenario", "P_raw", "P_counted", "P_excluded", "Q", "S", "S_arith", "PF", "StartDelta", "S_transient", "Rows",
            "ExclResolved", "Rank"]
    _sheet(wb, "Scenarios", keys, [[r[k] for k in keys] for r in sc])
    reg = res["application"]["register"]
    _sheet(wb, "Validation", ["ID", "Category", "Workbook result", "Application status", "Severity", "Recommended action", "Note"],
           [[r["id"], r["category"], r["source_status"], r["status"], r["severity"], r["action"], r["note"]] for r in reg])
    f = res["application"]["findings"]
    _sheet(wb, "Findings", ["Object", "Object ID", "Code", "Status", "Severity", "Description", "Recommended action", "Workbook check"],
           [[x["object"], x["object_id"], x["code"], x["status"], x["severity"], x["description"], x["action"], x["workbook_check"]] for x in f])
    _sheet(wb, "Unresolved", ["Category", "Item", "Reference", "Severity"],
           [[u["category"], u["item"], u["reference"], u["severity"]] for u in summ["unresolved"]])
    _sheet(wb, "Input Trace", ["Where", "Field", "Expression", "Value", "References", "Status"],
           [[t["where"], t["field"], t["expression"], t["value"], str(t["references"]), t["status"]] for t in res["input_trace"]])
    from services.json_export import LIMITATIONS
    _sheet(wb, "Limitations", ["Limitation"], [[l] for l in LIMITATIONS])
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def v11_workbook(project) -> bytes:
    """Write the project's inputs into a v1.1 formula workbook (formulas recalculate when opened in Excel)."""
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    sys.path.insert(0, os.path.join(root, "tools", "v11"))
    import v11_workbook_generator as G
    db = {k: v for k, v in (project.get("design_basis") or {}).items() if v is not None}
    cats = [(c.get("Category"), c.get("CatDescription")) for c in project.get("categories") or [] if c.get("Category")]
    lib = [dict(l) for l in project.get("library") or []]
    from calculations.constants import SC
    profs = [(p.get("Profile"), p.get("ProfDescription"), [p.get(s) for s in SC]) for p in project.get("profiles") or [] if p.get("Profile")]
    user = [v for v in (project.get("rating_series") or {}).get("User list", []) if isinstance(v, (int, float))]
    real = lambda rows: [r for r in rows or [] if not r.get("_padding")]
    cfg = dict(db=db, loads=real(project.get("loads")), tx=real(project.get("transformers")), fdr=real(project.get("feeders")),
               cats=cats, lib=lib, profiles=profs, user_std=user, days=(project.get("settings") or {}).get("days_per_year", 365),
               brk_scen=(project.get("settings") or {}).get("breakdown_scenario", "MAXSIM"),
               central_type=(project.get("settings") or {}).get("central_type", "Liquid-filled"))
    with tempfile.TemporaryDirectory() as d:
        p = os.path.join(d, "v11.xlsx")
        G.build(p, cfg)
        return open(p, "rb").read()
