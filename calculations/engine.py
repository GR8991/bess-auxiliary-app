"""Calculation orchestrator - evaluates the workbook v1.1 dependency graph in order
(docs/PHASE1_REVERSE_ENGINEERING_REPORT.md §3).  Pure, deterministic, no UI and no AI calls.

    results = run(project_dict)

`results["workbook"]` holds outputs that reproduce v1.1 cells (equivalence reference);
`results["application"]` holds application-level additions, each labelled NOT IN v1.1.
"""
from __future__ import annotations

from . import design_basis as DBm
from .excel_semantics import XLError
from .input_expressions import resolve_inputs
from .load_calculations import calculate_loads
from .transformer_sizing import active_std_series, tx_voltages, calculate_transformers, central_comparison
from .lv_distribution import calculate_feeders
from .scenario_analysis import site_results, summary
from .validation_engine import register, standards_applicability, library_status
from .constants import ENGINE_VERSION, SOURCE_WORKBOOK_VERSION
from . import reference_data as R
from . import app_layer


def _rows(project, key):
    """Excel Tables always hold at least one data row; an empty table is evaluated as one blank
    row, exactly as the v1.1 workbook does (affects e.g. V20 when no feeder is defined).
    Such rows carry '_padding': True so that the UI can hide them."""
    rows = [dict(r) for r in (project.get(key) or [])]
    return rows if rows else [{"_padding": True}]


def run(project: dict) -> dict:
    trace = []
    db = DBm.effective(project.get("design_basis") or {}, trace)

    def resolver(name):
        k = name[3:]
        if k not in db:
            raise XLError("#NAME?")
        return db[k]

    loads_in = [resolve_inputs(r, resolver, trace, f"loads[{i}]") for i, r in enumerate(_rows(project, "loads"))]
    tx_in = [resolve_inputs(r, resolver, trace, f"transformers[{i}]") for i, r in enumerate(_rows(project, "transformers"))]
    fdr_in = [resolve_inputs(r, resolver, trace, f"feeders[{i}]") for i, r in enumerate(_rows(project, "feeders"))]
    lib = [resolve_inputs(r, resolver, trace, f"library[{i}]") for i, r in enumerate(_rows(project, "library"))]
    cats = _rows(project, "categories")
    profs = _rows(project, "profiles")
    scen_rows = _rows(project, "scenarios")
    settings = dict(project.get("settings") or {})
    series = project.get("rating_series") or R.RATING_SERIES

    lst = active_std_series(db["StdSeries"], series)
    volts = tx_voltages(tx_in, db)
    ctx = dict(db=db, lib=lib, lib_ids=[l.get("LibID") for l in lib], profiles=profs,
               prof_ids=[p.get("Profile") for p in profs], cat_ids=[c.get("Category") for c in cats],
               tx_ids=[t.get("TX_ID") for t in tx_in], tx_vll=[v["VLL_e"] for v in volts], tx_vln=[v["VLN_e"] for v in volts],
               fdr_ids=[f.get("FeederID") for f in fdr_in], fdr_tx=[f.get("TX_ID") for f in fdr_in])
    loads = calculate_loads(loads_in, ctx)
    for row, r in zip(loads_in, loads):
        r["_TX"], r["_Feeder"], r["_PhaseAssign"] = row.get("TX"), row.get("Feeder"), row.get("PhaseAssign")
        r["_ExclGroup"], r["_Reference"] = row.get("ExclGroup"), row.get("Reference")
    scen = site_results(loads, scen_rows)
    tx_out = calculate_transformers(tx_in, volts, loads, db, lst)
    fdr_out = calculate_feeders(fdr_in, loads, tx_in, tx_out, db)
    dbchk = DBm.checks(db)
    from .excel_semantics import countif
    dbx = DBm.derived(db, dbchk, countif([t.get("TX_ID") for t in tx_in], "?*"))
    summ = summary(loads, scen, db, cats, tx_in, tx_out, settings)
    cmp_ = central_comparison(scen["peak_S"], tx_out, tx_in, db, lst, settings.get("central_type", "Liquid-filled"))
    val, counts = register(db, dbx, loads, scen, tx_in, tx_out, fdr_in, fdr_out, cats)
    wb = dict(design_basis=db, design_basis_checks=dbchk, design_basis_derived=dbx, std_active=lst,
              library_status=library_status(lib), loads=loads, scenarios=scen, transformers=tx_out, feeders=fdr_out,
              summary=summ, central_comparison=cmp_, validation=val, validation_counts=counts,
              standards_applicability=standards_applicability(db))
    res = dict(engine_version=ENGINE_VERSION, source_workbook_version=SOURCE_WORKBOOK_VERSION, input_trace=trace,
               inputs=dict(loads=loads_in, transformers=tx_in, feeders=fdr_in, library=lib, categories=cats, profiles=profs),
               workbook=wb)
    res["application"] = app_layer.build(res)
    return res
