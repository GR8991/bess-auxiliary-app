"""AI-Agent handoff package builder (brief §10). Machine-readable, schema-validated, every value tagged."""
from __future__ import annotations

import json
import math

from calculations import reference_data as R
from calculations.constants import ENGINE_VERSION, SOURCE_WORKBOOK_VERSION
from calculations.design_basis import default_design_basis
from calculations.excel_semantics import XLError
from calculations.load_calculations import CALC_COLUMNS
from calculations.transformer_sizing import TX_CALC_COLUMNS
from calculations.lv_distribution import FDR_CALC_COLUMNS
from data_models.handoff_schema import HANDOFF_SCHEMA, HANDOFF_SCHEMA_VERSION
from data_models.project_schema import APP_VERSION, SCHEMA_VERSION, PROJECT_INFO_FIELDS, DB_MANDATORY, now_iso
from services import engineering_summary as ES

LIMITATIONS = [
    "Results reproduce workbook v1.1 as recalculated by LibreOffice Calc; Microsoft Excel desktop recalculation not yet verified.",
    "Screening-level only: motor starting / inrush (F-05), voltage regulation, voltage drop and short-circuit values are approximations.",
    "R-1: regulation and voltage drop use the aggregate running PF, not the governing-scenario PF.",
    "R-2: transient kVA is site-wide (steady S + largest starting increment), one start at a time.",
    "R-3: starting increment uses rated S_unit; R-4: exclusivity ties go to the earliest row; R-5: ERR rows excluded from totals; "
    "R-6: loading % excludes design margins.",
    "R-7/E-1: empty tables are evaluated as one blank row (Excel behaviour); V20 lacks a 'no feeders' guard in v1.1.",
    "Linear ambient/altitude derating with user-entered coefficients; harmonic derating is a user input; no thermal loading-guide calculation.",
    "Contingency sizing assumes equal load sharing within a redundancy group.",
    "No standard compliance is claimed; editions and clauses require verification against licensed copies (F-07).",
    "A PASS indicates the workbook check is satisfied, not that the design is approved.",
]
LEGEND = {"user_input": "Entered by the engineer (or imported from the engineer's workbook input cells)",
          "workbook_default": "Value as delivered in the blank v1.1 template, not changed by the engineer",
          "derived": "Derived from other inputs by a workbook formula", "calculated": "Output of the deterministic calculation engine",
          "missing": "Required or expected input not provided - do NOT fill automatically"}


def clean(v):
    if isinstance(v, XLError):
        return v.code
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return None
    if isinstance(v, dict):
        return {k: clean(x) for k, x in v.items() if not str(k).startswith("_") or k in ("_extra",)}
    if isinstance(v, (list, tuple)):
        return [clean(x) for x in v]
    return v


def _db_tagged(project, res):
    dflt = default_design_basis()
    db_in = project.get("design_basis") or {}
    eff = res["workbook"]["design_basis"]
    chk = res["workbook"]["design_basis_checks"]
    out = {}
    for it in R.DB_ITEMS:
        k = it["key"]
        if it["derived"]:
            src = "derived"
        elif db_in.get(k) in (None, ""):
            src = "missing"
        elif dflt.get(k) is not None and db_in.get(k) == dflt.get(k):
            src = "workbook_default"
        else:
            src = "user_input"
        out[k] = dict(value=clean(eff.get(k)), source=src, unit=it["unit"], check=clean(chk.get("DB_" + k)),
                      expression=db_in.get(k) if isinstance(db_in.get(k), str) and str(db_in.get(k)).startswith("=") else None,
                      mandatory=k in DB_MANDATORY)
    return out


def build_handoff(project, res, generated_by=""):
    wb, app = res["workbook"], res["application"]
    summ = ES.build(project, res)
    meta = project.get("metadata") or {}
    info = project.get("project_info") or {}
    pinfo = {k: dict(value=info.get(k), source="user_input" if info.get(k) not in (None, "") else "missing", unit=None, check=None)
             for k, _ in PROJECT_INFO_FIELDS}
    real = lambda rows: [r for r in rows if not r.get("_padding")]
    loads_in, tx_in, fdr_in = real(res["inputs"]["loads"]), real(res["inputs"]["transformers"]), real(res["inputs"]["feeders"])
    raw_loads = real(project.get("loads") or [])
    load_calc = [dict({"ID": clean(r.get("ID"))}, **{c: clean(o[c]) for c in CALC_COLUMNS}) for r, o in zip(loads_in, wb["loads"])]
    tx_calc = [dict({"TX_ID": clean(t.get("TX_ID"))}, **{c: clean(o[c]) for c in TX_CALC_COLUMNS}) for t, o in zip(tx_in, wb["transformers"])]
    fdr_calc = [dict({"FeederID": clean(f.get("FeederID"))}, **{c: clean(o[c]) for c in FDR_CALC_COLUMNS}) for f, o in zip(fdr_in, wb["feeders"])]
    assumptions = [dict(clean(a), source="workbook v1.1 Assumptions register") for a in R.ASSUMPTIONS]
    recs = [dict(id=r["id"], category=r["category"], status=r["status"], severity=r["severity"], recommendation=r["action"])
            for r in app["register"] if r["status"] not in ("PASS", "NOT APPLICABLE")]
    pkg = dict(
        package_type="BESS_AUX_BOP_HANDOFF", handoff_schema_version=HANDOFF_SCHEMA_VERSION,
        identification=dict(project_id=meta.get("project_id"), project_name=wb["design_basis"].get("Project"), revision=meta.get("revision"),
                            generated_utc=now_iso(), generated_by=generated_by, application_version=APP_VERSION,
                            calculation_engine_version=ENGINE_VERSION, source_workbook_version=SOURCE_WORKBOOK_VERSION,
                            project_schema_version=SCHEMA_VERSION, handoff_schema_version=HANDOFF_SCHEMA_VERSION),
        project_information=pinfo,
        design_basis=_db_tagged(project, res),
        equipment_library=dict(source="user_input (ratings blank unless entered from datasheets)",
                               items=[dict(clean(l), status=clean(s)) for l, s in zip(real(project.get("library") or []), wb["library_status"])]),
        load_schedule=dict(user_inputs=clean(raw_loads), calculated=load_calc, input_formula_trace=clean(res["input_trace"])),
        operating_scenarios=dict(definitions=clean(project.get("scenarios")), profiles=clean(project.get("profiles")),
                                 profiles_reviewed=wb["design_basis"].get("ProfRev"),
                                 results=clean(wb["scenarios"]["rows"]), peak_scenario=wb["scenarios"]["peak_name"],
                                 peak_kVA=wb["scenarios"]["peak_S"]),
        transformer_sizing=dict(user_inputs=clean(real(project.get("transformers") or [])), calculated=tx_calc,
                                centralized_comparison=clean(wb["central_comparison"]), summary_table=clean(summ["transformers"]),
                                note="Req_min = calculated minimum; Rec_std = recommended standard rating (NOT approved); "
                                     "Selected_kVA = engineer-selected rating."),
        lv_distribution=dict(user_inputs=clean(real(project.get("feeders") or [])), calculated=fdr_calc,
                             application_feeder_phase_balance=clean(app["feeder_phase_balance"]), note="Screening checks only."),
        assumptions=assumptions,
        validation=dict(register=clean(app["register"]), findings=clean(app["findings"]), status_counts=app["status_counts"],
                        workbook_counts=wb["validation_counts"]),
        outstanding_data_requirements=summ["unresolved"],
        engineering_limitations=LIMITATIONS,
        engineering_recommendations=recs,
        methodology=dict(workbook="BESS Auxiliary Load & Transformer Sizing v1.1", engine_version=ENGINE_VERSION,
                         equivalence="119,876 comparisons vs LibreOffice-recalculated v1.1 workbook, 0 fail (see docs/CALCULATION_EQUIVALENCE_REPORT.md)",
                         formula_mapping="docs/formula_mapping_v1.1.csv", application_level_additions=app["label"],
                         llm_used_for_calculation=False),
        input_completion=summ["completion"],
        overall_disposition=dict(status=app["disposition"], note=summ["disposition_note"]),
        electrical_summary=clean(summ["electrical"]),
        value_source_legend=LEGEND,
    )
    return clean(pkg)


def validate_handoff(pkg):
    import jsonschema
    jsonschema.validate(pkg, HANDOFF_SCHEMA)
    return True


def to_bytes(pkg):
    return json.dumps(pkg, indent=2, ensure_ascii=False, default=str).encode("utf-8")
