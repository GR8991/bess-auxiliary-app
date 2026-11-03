"""Sheet 3 - Load Schedule (tblLoad): all 58 calculated columns of the v1.1 workbook.

Each function reproduces one workbook column formula; the comment on each line gives the Excel
column.  Rows are processed in table order (row order matters for the exclusivity tie-break, R-4).
"""
from __future__ import annotations

import math

from .excel_semantics import (SQRT3, XLError, isnumber, isblank_text, N, num, eq, ne, gt, lt, le, ge, left, text,
                              countifs, match_exact, safe)
from .constants import SC

LOAD_INPUTS = ['ID', 'LibID', 'Description', 'Category', 'TX', 'Feeder', 'QtyInst', 'QtyDuty', 'Phase', 'Conn1Ph',
               'PhaseAssign', 'VOverride', 'RatingType', 'Rating', 'PF', 'Reactive', 'Eff', 'DF', 'DutyF', 'Hours', 'Profile',
               'ExclGroup', 'SupplyClass', 'StartMult', 'StartPF', 'Nonlinear', 'THDi', 'Reference', 'Remarks']


def dbref(v):
    """A formula that returns a reference to a Design-Basis cell displays 0 when that cell is empty."""
    return 0 if v is None else v


def _lib_text(lib, col, idx):
    return text(lib[idx - 1].get(col))


def _lib_num(lib, col, idx):
    v = lib[idx - 1].get(col)
    return v if isnumber(v) else ""


def _num_eff(row, col, lib, libcol, idx):
    v = row.get(col)
    if isnumber(v):
        return v
    if idx > 0:
        return _lib_num(lib, libcol, idx)
    return ""


def _text_eff(row, col, lib, libcol, idx):
    v = row.get(col)
    if ne(v, ""):
        return v
    if idx > 0:
        return _lib_text(lib, libcol, idx)
    return ""


def row_pass1(i, row, ctx):
    """Columns that depend only on this row's inputs, library, profiles, DB and TX/feeder inputs."""
    lib, db = ctx["lib"], ctx["db"]
    r = {}
    r["Active"] = 0 if (isblank_text(row.get("ID")) and isblank_text(row.get("LibID")) and isblank_text(row.get("Description"))
                        and eq(row.get("Rating"), "")) else 1
    r["Row"] = i + 1
    if eq(row.get("LibID"), ""):
        r["LibIdx"] = 0
    else:
        m = match_exact(row.get("LibID"), ctx["lib_ids"])
        r["LibIdx"] = m if m else -1
    idx = r["LibIdx"]
    r["Cat_e"] = _text_eff(row, "Category", lib, "Category", idx)
    r["Phase_e"] = _text_eff(row, "Phase", lib, "Phase", idx)
    r["RType_e"] = _text_eff(row, "RatingType", lib, "RatingType", idx)
    r["Rating_e"] = _num_eff(row, "Rating", lib, "RatedValue", idx)
    r["PF_e"] = 1 if eq(r["Phase_e"], "DC") else _num_eff(row, "PF", lib, "PF", idx)
    r["Eff_e"] = _num_eff(row, "Eff", lib, "Eff", idx)
    r["StartM_e"] = _num_eff(row, "StartMult", lib, "StartMult", idx)
    r["StartPF_e"] = _num_eff(row, "StartPF", lib, "StartPF", idx)
    r["NL_e"] = _text_eff(row, "Nonlinear", lib, "Nonlinear", idx)
    r["Prof_e"] = _text_eff(row, "Profile", lib, "Profile", idx)
    q = row.get("QtyDuty")
    r["QtyDuty_e"] = q if isnumber(q) else N(row.get("QtyInst"))
    r["DF_e"] = row["DF"] if isnumber(row.get("DF")) else dbref(db["DefDF"])
    r["DutyF_e"] = row["DutyF"] if isnumber(row.get("DutyF")) else dbref(db["DefDuty"])
    r["Hours_e"] = row["Hours"] if isnumber(row.get("Hours")) else dbref(db["DefHours"])
    tx = row.get("TX")
    if eq(tx, ""):
        r["VLL"], r["VLN"] = N(db["VLL"]), N(db["VLN"])
    else:
        m = match_exact(tx, ctx["tx_ids"])
        r["VLL"] = ctx["tx_vll"][m - 1] if m else N(db["VLL"])
        r["VLN"] = ctx["tx_vln"][m - 1] if m else N(db["VLN"])
    vo = row.get("VOverride")
    if isnumber(vo):
        r["V_app"] = vo
    elif eq(r["Phase_e"], "3PH"):
        r["V_app"] = r["VLL"]
    elif eq(r["Phase_e"], "1PH"):
        r["V_app"] = r["VLL"] if eq(row.get("Conn1Ph"), "L-L") else r["VLN"]
    elif eq(r["Phase_e"], "DC"):
        r["V_app"] = N(db["VDC"])
    else:
        r["V_app"] = 0
    m = match_exact(r["Prof_e"], ctx["prof_ids"])
    r["PfIdx"] = m if m else 0
    return r


def row_check(row, r, ctx, all_rows):
    """Load-row Check: first issue of the 27-level workbook chain (string identical to Excel)."""
    if r["Active"] == 0:
        return ""
    if r["LibIdx"] == -1:
        return "ERR: LibID not in Equipment Library"
    ph, rt = r["Phase_e"], r["RType_e"]
    if eq(ph, ""):
        return "ERR: phase missing"
    if ne(ph, "3PH") and ne(ph, "1PH") and ne(ph, "DC"):
        return "ERR: invalid phase"
    if ne(rt, "kW_in") and ne(rt, "kVA") and ne(rt, "A") and ne(rt, "kW_out"):
        return "ERR: rating type missing/invalid"
    if not isnumber(r["Rating_e"]):
        return "ERR: rating missing (row or library)"
    if lt(r["Rating_e"], 0):
        return "ERR: negative rating"
    pf = r["PF_e"]
    if (not isnumber(pf)) or le(pf, 0) or gt(pf, 1):
        return "ERR: PF must satisfy 0 < PF <= 1"
    eff = r["Eff_e"]
    if eq(rt, "kW_out") and ((not isnumber(eff)) or le(eff, 0) or gt(eff, 1)):
        return "ERR: efficiency 0<η<=1 required for kW_out"
    qi = row.get("QtyInst")
    if (not isnumber(qi)) or le(qi, 0):
        return "ERR: installed quantity must be > 0"
    if lt(r["QtyDuty_e"], 0) or gt(r["QtyDuty_e"], qi):
        return "ERR: duty qty must be 0..installed qty"
    if lt(r["DF_e"], 0) or gt(r["DF_e"], 1) or lt(r["DutyF_e"], 0) or gt(r["DutyF_e"], 1):
        return "ERR: DF / duty factor outside 0-1"
    if lt(r["Hours_e"], 0) or gt(r["Hours_e"], 24):
        return "ERR: hours/day outside 0-24"
    if r["PfIdx"] == 0:
        return "ERR: scenario profile missing/invalid"
    tx, fd = row.get("TX"), row.get("Feeder")
    if ne(tx, "") and match_exact(tx, ctx["tx_ids"]) is None:
        return "ERR: transformer not defined on sheet 6"
    if ne(fd, "") and match_exact(fd, ctx["fdr_ids"]) is None:
        return "ERR: feeder not defined on sheet 7"
    if ne(fd, "") and ne(tx, ""):
        m = match_exact(fd, ctx["fdr_ids"])
        ftx = ctx["fdr_tx"][m - 1] if m else tx
        if ne(ftx, tx):
            return "ERR: feeder is supplied from a different transformer"
    va = r["V_app"]
    if (not isnumber(va)) or le(va, 0):
        return "ERR: voltage undefined (set LV/DC voltage or override)"
    rc = row.get("Reactive")
    if ne(rc, "") and ne(rc, "Lag") and ne(rc, "Lead"):
        return "ERR: Reactive must be Lag or Lead"
    g = row.get("ExclGroup")
    if ne(g, ""):
        grp = [x.get("ExclGroup") for x in all_rows]
        txs = [x.get("TX") for x in all_rows]
        fds = [x.get("Feeder") for x in all_rows]
        n = countifs(grp, g, txs, "<>" + text(tx)) + countifs(grp, g, fds, "<>" + text(fd))
        if n > 0:
            return "ERR: ExclGroup members must share the same TX and Feeder"
    if eq(tx, "") and ne(ph, "DC"):
        return "WARN: no transformer assigned (site totals only)"
    if eq(r["Cat_e"], ""):
        return "WARN: category missing"
    if match_exact(r["Cat_e"], ctx["cat_ids"]) is None:
        return "WARN: category not in tblCat"
    if eq(ph, "1PH") and eq(row.get("Conn1Ph"), ""):
        return "WARN: 1PH connection blank - L-N assumed"
    if eq(ph, "DC"):
        return "INFO: DC load - excluded from AC totals (supply via charger row)"
    if (not isnumber(row.get("DF"))) or (not isnumber(row.get("DutyF"))) or (not isnumber(row.get("Hours"))):
        return "INFO: design-basis default DF/duty/hours applied"
    if eq(row.get("Reference"), ""):
        return "INFO: no datasheet reference"
    return "OK"


def row_pass2(row, r, ctx):
    """Electrical quantities (after Check/Valid)."""
    prof = ctx["profiles"]
    r["Valid"] = 1 if (r["Active"] == 1 and left(r["Check"], 3) != "ERR") else 0
    v, rt, ph = r["Valid"], r["RType_e"], r["Phase_e"]
    R, pf, V = r["Rating_e"], r["PF_e"], r["V_app"]
    if v == 0:
        S = 0
    elif eq(rt, "kVA"):
        S = R
    elif eq(rt, "A"):
        S = SQRT3 * V * R / 1000 if eq(ph, "3PH") else V * R / 1000
    elif eq(rt, "kW_in"):
        S = R / pf
    elif eq(rt, "kW_out"):
        S = R / r["Eff_e"] / pf
    else:
        S = 0
    r["S_unit"] = S
    r["P_unit"] = 0 if v == 0 else S * pf
    r["I_unit"] = 0 if v == 0 else (S * 1000 / (SQRT3 * V) if eq(ph, "3PH") else S * 1000 / V)
    r["TanPhi"] = 0 if v == 0 else (-1 if eq(row.get("Reactive"), "Lead") else 1) * math.sqrt(1 - pf ** 2) / pf
    r["IsAC"] = 1 if (v == 1 and ne(ph, "DC")) else 0
    qi = N(row.get("QtyInst"))
    r["Conn_kW"] = r["IsAC"] * qi * r["P_unit"]
    r["Conn_kVA"] = r["IsAC"] * qi * r["S_unit"]
    r["Run_kW"] = r["IsAC"] * r["QtyDuty_e"] * r["P_unit"] * r["DF_e"]
    r["Run_kVA"] = r["IsAC"] * r["QtyDuty_e"] * r["S_unit"] * r["DF_e"]
    r["DC_kW"] = r["QtyDuty_e"] * r["P_unit"] * r["DF_e"] if (v == 1 and eq(ph, "DC")) else 0
    r["Start_kVA"] = r["IsAC"] * r["S_unit"] * (r["StartM_e"] if isnumber(r["StartM_e"]) else 1)
    r["Start_kW"] = r["Start_kVA"] * (r["StartPF_e"] if isnumber(r["StartPF_e"]) else N(r["PF_e"]))
    r["StartDelta"] = max(0, r["Start_kVA"] - r["IsAC"] * r["S_unit"])        # R-3: rated S, not DF-scaled
    r["Day_kWh"] = r["Run_kW"] * r["DutyF_e"] * r["Hours_e"]
    for s in SC:
        f = prof[r["PfIdx"] - 1].get(s) if r["PfIdx"] else None
        r["R_" + s] = 0 if r["PfIdx"] == 0 else r["Run_kW"] * num(f)
    for s in SC:
        r["RS_" + s] = r["R_" + s] * math.sqrt(1 + r["TanPhi"] ** 2)
    return r


def apply_exclusivity(rows_in, res):
    """P_<s> counted kW (F-01): within an ExclGroup only the earliest row having the group's maximum
    raw kVA (RS_<s>) is counted, with its own P (and therefore its own Q = P*tanφ)."""
    groups = [x.get("ExclGroup") for x in rows_in]
    for s in SC:
        RS = [r["RS_" + s] for r in res]
        for i, r in enumerate(res):
            g = groups[i]
            if eq(g, ""):
                r["P_" + s] = r["R_" + s]
                continue
            member = [eq(gg, g) for gg in groups]
            gmax = max((RS[j] if member[j] else 0) for j in range(len(res)))
            earlier = sum(1 for j in range(len(res)) if member[j] and RS[j] == RS[i] and res[j]["Row"] < r["Row"])
            r["P_" + s] = r["R_" + s] if (RS[i] == gmax and earlier == 0) else 0


CALC_COLUMNS = (["Check", "Active", "Row", "LibIdx", "Cat_e", "Phase_e", "RType_e", "Rating_e", "PF_e", "Eff_e", "StartM_e",
                 "StartPF_e", "NL_e", "Prof_e", "QtyDuty_e", "DF_e", "DutyF_e", "Hours_e", "VLL", "VLN", "V_app", "PfIdx", "Valid",
                 "S_unit", "P_unit", "I_unit", "TanPhi", "IsAC", "Conn_kW", "Conn_kVA", "Run_kW", "Run_kVA", "DC_kW", "Start_kVA",
                 "Start_kW", "StartDelta", "Day_kWh"] + [f"R_{s}" for s in SC] + [f"RS_{s}" for s in SC] + [f"P_{s}" for s in SC])


APP_INPUT_ERROR = "ERR: input formula could not be evaluated (application-level, see input trace)"


def calculate_loads(rows_in, ctx):
    """Excel would propagate #VALUE! from an erroneous input formula into every dependent cell and
    total.  The application instead marks that row with APP_INPUT_ERROR (invalid -> excluded, flagged)
    so that the rest of the project remains calculable.  This path never occurs for inputs that
    the workbook itself evaluates without error (equivalence datasets)."""
    res = []
    for i, row in enumerate(rows_in):
        bad = any(isinstance(v, XLError) for v in row.values())
        clean = {k: (None if isinstance(v, XLError) else v) for k, v in row.items()}
        r = row_pass1(i, clean, ctx)
        r["_input_error"] = bad
        res.append(r)
    for row, r in zip(rows_in, res):
        clean = {k: (None if isinstance(v, XLError) else v) for k, v in row.items()}
        r["Check"] = APP_INPUT_ERROR if (r["_input_error"] and r["Active"] == 1) else row_check(clean, r, ctx, rows_in)
        row_pass2(clean, r, ctx)
    apply_exclusivity(rows_in, res)
    return res
