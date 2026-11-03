"""Sheet 1 - Design Basis: effective values, derived values and column-F input checks.
Every check reproduces the v1.1 cell formula (see docs/formula_mapping_v1.1.csv, EQ-DB-*)."""
from __future__ import annotations

from .excel_semantics import SQRT3, isnumber, isblank_text, gt, ge, le, eq, countif, safe, XLError
from .input_expressions import is_expression, evaluate, UnsupportedExpression
from . import reference_data as R

DB_KEYS = [i["key"] for i in R.DB_ITEMS]
DERIVED = {"VLN"}
REQUIRED = {"Jurisdiction", "Arrangement", "Tamb", "Alt"}


def default_design_basis() -> dict:
    """Values exactly as delivered in the blank v1.1 template (blank = None)."""
    return {i["key"]: i["default"] for i in R.DB_ITEMS if not i["derived"]}


def effective(db_in: dict, trace: list) -> dict:
    """Resolve input formulas (A-6) and derived values. Unknown keys are preserved untouched."""
    raw = {k: db_in.get(k) for k in DB_KEYS if k not in DERIVED}
    resolved, stack = {}, []

    def get(k):
        if k in resolved:
            return resolved[k]
        if k in stack:
            raise UnsupportedExpression(f"circular reference via DB_{k}")
        v = raw.get(k) if k != "VLN" else None
        if k == "VLN":
            v = vln()
        elif is_expression(v):
            stack.append(k)
            try:
                val, refs = evaluate(v, lambda name: get(name[3:]))
                trace.append(dict(where="design_basis", field=k, expression=v, value=val, references=refs, status="OK"))
                v = val
            except XLError as e:
                trace.append(dict(where="design_basis", field=k, expression=v, value=e.code, references={}, status="EXCEL ERROR"))
                v = e
            except UnsupportedExpression as e:
                trace.append(dict(where="design_basis", field=k, expression=v, value=None, references={}, status=f"UNSUPPORTED: {e}"))
                v = XLError("#UNSUPPORTED")
            stack.pop()
        resolved[k] = v
        return v

    def vln():
        o, l = get("VLN_ovr"), get("VLL")
        if isnumber(o):
            return o
        if isnumber(l):
            return l / SQRT3
        return ""

    for k in DB_KEYS:
        get(k)
    return resolved


def _check(key, c):
    if key in ("Jurisdiction", "Arrangement", "Tamb", "Alt"):
        return "ERR: required input missing" if isblank_text(c) else "OK"
    if key == "Freq":
        if isblank_text(c):
            return "ERR: required input missing"
        return "OK" if (eq(c, 50) or eq(c, 60)) else "ERR: 50 or 60 Hz only"
    if key in ("MV", "VLL"):
        if isblank_text(c):
            return "ERR: required input missing"
        return "OK" if (isnumber(c) and gt(c, 0)) else "ERR: must be a number > 0"
    if key in ("VLN_ovr", "VDC"):
        if isblank_text(c):
            return "OK (blank / not used)"
        return "OK" if (isnumber(c) and gt(c, 0)) else "ERR: must be a number > 0"
    if key in ("DerK_L", "DerAlt_L", "DerK_D", "DerAlt_D", "RegLim", "DipLim", "VDLim", "NLThr", "ImbLim", "SrcMVA", "Cmax"):
        if isblank_text(c):
            return "REVIEW: not entered"
        return "OK" if (isnumber(c) and ge(c, 0)) else "ERR: must be a number >= 0"
    if key in ("Margin", "Future"):
        if isblank_text(c):
            return "REVIEW: blank - 0 % applied"
        return "OK" if (isnumber(c) and ge(c, 0) and le(c, 1)) else "ERR: 0-100 %"
    if key in ("MaxLoad", "MaxLoadCont"):
        return "OK" if (isnumber(c) and gt(c, 0)) else "ERR: > 0 required"
    if key in ("DefDF", "DefDuty"):
        return "OK" if (isnumber(c) and ge(c, 0) and le(c, 1)) else "ERR: 0 to 1 required"
    if key == "DefHours":
        return "OK" if (isnumber(c) and ge(c, 0) and le(c, 24)) else "ERR: 0 to 24 h"
    if key in ("ProfRev", "StdRev"):
        return "OK" if eq(c, "Yes") else "REVIEW: not yet confirmed"
    return None  # no check cell in workbook


def checks(db: dict) -> dict:
    """Column F of Design Basis, keyed 'DB_<key>' (None where the workbook has no check)."""
    out = {}
    for k in DB_KEYS:
        r = safe(_check, k, db[k])
        if r is not None:
            out["DB_" + k] = r
    return out


def derived(db: dict, chk: dict, n_tx: int) -> dict:
    vals = list(chk.values())
    mw, mwh = db["BESS_MW"], db["BESS_MWh"]
    dur = mwh / mw if (isnumber(mw) and mw > 0 and isnumber(mwh)) else ""
    return dict(Duration_h=dur, N_TX_defined=n_tx, DB_ErrCount=countif(vals, "ERR*"), DB_RevCount=countif(vals, "REVIEW*"))
