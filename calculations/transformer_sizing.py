"""Sheet 6 - Transformer Sizing (tblTX, 69 calculated columns) + centralized comparison block,
and Lists!lstStdActive.  Reproduces v1.1 formulas exactly (including R-1 aggregate running PF
for regulation and the F-04 DerateOK gate)."""
from __future__ import annotations

import math

from .excel_semantics import (SQRT3, XLError, isnumber, isblank_text, N, num, eq, ne, gt, lt, le, left, text,
                              sumifs, countif, countifs, match_exact, xmax, count, small, num_to_text, safe)
from .constants import SC
from .load_calculations import dbref

TX_INPUTS = ['TX_ID', 'Description', 'Section', 'Role', 'RedGroup', 'Scheme', 'HV_kV', 'LV_VLL', 'LV_VLN', 'Type', 'Cooling',
             'Z_pct', 'XR', 'Kharm_in', 'InrushMult', 'DesignScen', 'Selected_kVA', 'Notes']
EPS = 0.000001  # the workbook's literal comparison allowance (+0.000001 / -0.000001)


def active_std_series(series_name, series: dict, n=40):
    """Lists!lstStdActive (40 rows): IEEE / User / else IEC; non-numbers -> ""."""
    def col(name):
        v = list(series.get(name, []))
        return (v + [None] * n)[:n]
    a, b, c = col("IEC R10"), col("IEEE 3-phase"), col("User list")
    out = []
    for i in range(n):
        if eq(series_name, "IEEE 3-phase"):
            out.append(b[i] if isnumber(b[i]) else "")
        elif eq(series_name, "User list"):
            out.append(c[i] if isnumber(c[i]) else "")
        else:
            out.append(a[i] if isnumber(a[i]) else "")
    return out


def tx_voltages(tx_rows, db):
    out = []
    for t in tx_rows:
        hv = t.get("HV_kV")
        vll = t["LV_VLL"] if isnumber(t.get("LV_VLL")) else N(db["VLL"])
        vln = t["LV_VLN"] if isnumber(t.get("LV_VLN")) else vll / SQRT3
        out.append(dict(HVkV_e=hv if isnumber(hv) else N(db["MV"]), VLL_e=vll, VLN_e=vln))
    return out


def _derate_parts(typ, db):
    dry = eq(typ, "Dry-type")
    refT, derK = (db["RefT_D"], db["DerK_D"]) if dry else (db["RefT_L"], db["DerK_L"])
    refA, derA = (db["RefAlt_D"], db["DerAlt_D"]) if dry else (db["RefAlt_L"], db["DerAlt_L"])
    return refT, derK, refA, derA


def k_ambient(typ, db):
    refT, derK, _, _ = _derate_parts(typ, db)
    if gt(N(db["Tamb"]), refT):
        return max(0, 1 - (num(db["Tamb"]) - num(refT)) * derK / 100) if isnumber(derK) else 1
    return 1


def k_altitude(typ, db):
    _, _, refA, derA = _derate_parts(typ, db)
    if gt(N(db["Alt"]), refA):
        return max(0, 1 - (num(db["Alt"]) - num(refA)) / 100 * derA / 100) if isnumber(derA) else 1
    return 1


def recommended_standard(req, lst, with_count_check=True):
    if le(req, 0):
        return "No load"
    if with_count_check and count(lst) == 0:
        return "No std list"
    if gt(req, xmax(lst)):
        return "Exceeds std list"
    return small(lst, countif(lst, "<" + num_to_text(req)) + 1)


def _one_tx(t, v, loads, db):
    o = dict(v)
    tid = t.get("TX_ID")
    blank = eq(tid, "")
    L = loads
    ltx = [r["_TX"] for r in L]
    o["Scen_e"] = t["DesignScen"] if ne(t.get("DesignScen"), "") else dbref(db["DesignScen"])
    mask = [eq(x, tid) for x in ltx]
    for s in SC:
        if blank:
            o["S_" + s] = 0
        else:
            P = sumifs([r["P_" + s] for r in L], ltx, tid)
            Q = sum(r["P_" + s] * r["TanPhi"] for r, m in zip(L, mask) if m)
            o["S_" + s] = math.sqrt(P ** 2 + Q ** 2)
    o["Conn_kVA"] = 0 if blank else sumifs([r["Conn_kVA"] for r in L], ltx, tid)
    o["Conn_kW"] = 0 if blank else sumifs([r["Conn_kW"] for r in L], ltx, tid)
    o["P_run"] = 0 if blank else sumifs([r["Run_kW"] for r in L], ltx, tid)
    o["Q_run"] = 0 if blank else sum(r["Run_kW"] * r["TanPhi"] for r, m in zip(L, mask) if m)
    P, Q = o["P_run"], o["Q_run"]
    o["PF_run"] = P / math.sqrt(P ** 2 + Q ** 2) if P > 0 else ""
    o["Sin_run"] = Q / math.sqrt(P ** 2 + Q ** 2) if P > 0 else 0
    Ss = [o["S_" + s] for s in SC]
    if eq(o["Scen_e"], "ENVELOPE"):
        o["Design_kVA"] = max(Ss)
    else:
        m = match_exact(o["Scen_e"], SC)
        o["Design_kVA"] = Ss[m - 1] if m else 0
    d = o["Design_kVA"]
    if d == 0:
        o["Gov"] = "-"
    elif ne(o["Scen_e"], "ENVELOPE"):
        o["Gov"] = o["Scen_e"]
    else:
        o["Gov"] = next((s for s in SC if o["S_" + s] == d), "-")
    o["Kamb"] = k_ambient(t.get("Type"), db)
    o["Kalt"] = k_altitude(t.get("Type"), db)
    o["Kharm"] = t["Kharm_in"] if isnumber(t.get("Kharm_in")) else 1
    o["Kderate"] = o["Kamb"] * o["Kalt"] * o["Kharm"]
    o["DesignPlus"] = d * (1 + N(db["Margin"])) * (1 + N(db["Future"]))
    o["Req_own"] = o["DesignPlus"] / o["Kderate"] if o["Kderate"] > 0 else 0
    o["Scheme_e"] = "N" if eq(t.get("Scheme"), "") else t["Scheme"]
    if blank:
        o["NL_share"] = 0
    else:
        tot = sumifs([r["Run_kVA"] for r in L], ltx, tid)
        o["NL_share"] = sumifs([r["Run_kVA"] for r in L], ltx, tid, [r["NL_e"] for r in L], "Y") / tot if tot > 0 else 0
    refT, derK, refA, derA = _derate_parts(t.get("Type"), db)
    if blank:
        o["DerateOK"] = 1
    else:
        bad = ((gt(N(db["Tamb"]), refT) and not isnumber(derK)) or (gt(N(db["Alt"]), refA) and not isnumber(derA))
               or (gt(o["NL_share"], N(db["NLThr"])) and not isnumber(t.get("Kharm_in"))))
        o["DerateOK"] = 0 if bad else 1
    o["MaxStart_kVA"] = 0 if blank else max([r["Start_kVA"] if m else 0 for r, m in zip(L, mask)] or [0])
    # phase balance (F-03)
    if blank:
        for k in ("L1_kVA", "L2_kVA", "L3_kVA", "Unalloc_kVA"):
            o[k] = 0
    else:
        rk = [r["Run_kVA"] for r in L]
        ph = [r["Phase_e"] for r in L]
        pa = [r["_PhaseAssign"] for r in L]
        three = sumifs(rk, ltx, tid, ph, "3PH") / 3
        one = lambda p: sumifs(rk, ltx, tid, ph, "1PH", pa, p)
        o["L1_kVA"] = three + one("L1") + (one("L1-L2") + one("L3-L1")) / 2
        o["L2_kVA"] = three + one("L2") + (one("L1-L2") + one("L2-L3")) / 2
        o["L3_kVA"] = three + one("L3") + (one("L2-L3") + one("L3-L1")) / 2
        o["Unalloc_kVA"] = one("")
    tot3 = o["L1_kVA"] + o["L2_kVA"] + o["L3_kVA"]
    o["Imb_pct"] = (max(o["L1_kVA"], o["L2_kVA"], o["L3_kVA"]) - tot3 / 3) / (tot3 / 3) * 100 if tot3 > 0 else 0
    if blank:
        o["BalCheck"] = ""
    elif o["Unalloc_kVA"] > 0:
        o["BalCheck"] = "REVIEW: 1PH loads without phase allocation"
    elif not isnumber(db["ImbLim"]):
        o["BalCheck"] = "REVIEW: limit not set"
    else:
        o["BalCheck"] = "PASS" if o["Imb_pct"] <= db["ImbLim"] * 100 else "FAIL"
    return o


def _group_pass(tx_rows, out, db, lst):
    ids = [t.get("TX_ID") for t in tx_rows]
    rg = [t.get("RedGroup") for t in tx_rows]
    sch = [o["Scheme_e"] for o in out]
    des = [o["Design_kVA"] for o in out]
    plus = [o["DesignPlus"] for o in out]
    for t, o in zip(tx_rows, out):
        g = t.get("RedGroup")
        o["GrpN"] = 1 if eq(g, "") else countif(rg, g)
        o["GrpDesign"] = o["Design_kVA"] if eq(g, "") else sumifs(des, rg, g)
        o["GrpPlus"] = o["DesignPlus"] if eq(g, "") else sumifs(plus, rg, g)
        s = o["Scheme_e"]
        o["UnitsAvail"] = o["GrpN"] - 1 if eq(s, "N+1") else (o["GrpN"] / 2 if eq(s, "2N") else o["GrpN"])
        if eq(s, "N"):
            o["Req_cont"] = 0
        else:
            o["Req_cont"] = o["GrpPlus"] / o["UnitsAvail"] / o["Kderate"] if (o["UnitsAvail"] > 0 and o["Kderate"] > 0) else 0
        o["Req_min"] = max(o["Req_own"], o["Req_cont"])
        o["Rec_std"] = recommended_standard(o["Req_min"], lst)
        sel = t.get("Selected_kVA")
        o["Eval_kVA"] = sel if isnumber(sel) else (o["Rec_std"] if isnumber(o["Rec_std"]) else 0)
        if eq(t.get("TX_ID"), ""):
            o["RatingStatus"] = ""
        elif isnumber(sel):
            if sel < o["Req_min"] - EPS:
                o["RatingStatus"] = "FAIL: SELECTED below calculated minimum"
            elif o["DerateOK"] == 0:
                o["RatingStatus"] = "REVIEW: selected - derating data incomplete, sizing not confirmed"
            else:
                o["RatingStatus"] = "SELECTED - adequate"
        elif isnumber(o["Rec_std"]):
            o["RatingStatus"] = "PROVISIONAL - derating data incomplete" if o["DerateOK"] == 0 else "PROVISIONAL - recommended std, not approved"
        else:
            o["RatingStatus"] = "REVIEW: no rating available"
        ev, K, d = o["Eval_kVA"], o["Kderate"], o["Design_kVA"]
        o["Cap_derated"] = ev * K
        cap = o["Cap_derated"]
        o["Load_pct"] = d / ev * 100 if ev > 0 else 0                     # R-6: demand without margins
        o["LoadDer_pct"] = d / cap * 100 if cap > 0 else 0
        o["Spare_kVA"] = "" if eq(t.get("TX_ID"), "") else cap - d
        o["Spare_pct"] = o["Spare_kVA"] / cap * 100 if N(cap) > 0 else 0
        if eq(s, "N"):
            o["ContLoad_pct"] = ""
        else:
            o["ContLoad_pct"] = o["GrpDesign"] / o["UnitsAvail"] / cap * 100 if (o["UnitsAvail"] > 0 and cap > 0) else ""
        for sc in SC:
            o["U_" + sc] = o["S_" + sc] / ev * 100 if ev > 0 else 0
        hv, vll = o["HVkV_e"], o["VLL_e"]
        o["I_HV"] = ev / (SQRT3 * hv) if (ev > 0 and gt(hv, 0)) else 0
        o["I_LV"] = ev * 1000 / (SQRT3 * vll) if (ev > 0 and gt(vll, 0)) else 0
        o["I_LV_design"] = d * 1000 / (SQRT3 * vll) if gt(vll, 0) else 0
        z, xr = t.get("Z_pct"), t.get("XR")
        o["ur"] = z / math.sqrt(1 + xr ** 2) if (isnumber(z) and isnumber(xr)) else ""
        o["ux"] = o["ur"] * xr if isnumber(o["ur"]) else ""
        if isnumber(o["ur"]) and ev > 0 and isnumber(o["PF_run"]):          # R-1: aggregate running PF
            n_, pf, sn = d / ev, o["PF_run"], o["Sin_run"]
            o["Reg_pct"] = n_ * (o["ur"] * pf + o["ux"] * sn) + (n_ * (o["ux"] * pf - o["ur"] * sn)) ** 2 / 200
        else:
            o["Reg_pct"] = ""
        o["Ssc_tx"] = ev / (z / 100) if (isnumber(z) and ev > 0) else ""
        if isnumber(o["Ssc_tx"]):
            o["Ssc_bus"] = 1 / (1 / o["Ssc_tx"] + 1 / (db["SrcMVA"] * 1000)) if N(db["SrcMVA"]) > 0 else o["Ssc_tx"]
        else:
            o["Ssc_bus"] = ""
        o["Isc_kA"] = db["Cmax"] * o["Ssc_bus"] / (SQRT3 * vll) if (isnumber(o["Ssc_bus"]) and isnumber(db["Cmax"]) and gt(vll, 0)) else ""
        o["Dip_pct"] = o["MaxStart_kVA"] / (o["MaxStart_kVA"] + o["Ssc_bus"]) * 100 if (isnumber(o["Ssc_bus"]) and o["MaxStart_kVA"] > 0) else ""
        im = t.get("InrushMult")
        o["Inrush_A"] = im * o["I_HV"] if isnumber(im) else ""
    for t, o in zip(tx_rows, out):
        o["Check"] = safe(_tx_check, t, o, db, ids, rg, sch)


def _tx_check(t, o, db, ids, rg, sch):
    tid = t.get("TX_ID")
    if eq(tid, ""):
        return ""
    if countif(ids, tid) > 1:
        return "ERR: duplicate TX_ID"
    if le(o["HVkV_e"], 0):
        return "ERR: HV voltage missing"
    if le(o["VLL_e"], 0):
        return "ERR: LV voltage missing"
    typ = t.get("Type")
    if ne(typ, "Liquid-filled") and ne(typ, "Dry-type"):
        return "ERR: transformer type missing"
    s, g = o["Scheme_e"], t.get("RedGroup")
    if eq(s, "N+1") and o["GrpN"] < 2:
        return "ERR: N+1 needs >=2 units in RedGroup"
    if eq(s, "2N") and (o["GrpN"] < 2 or o["GrpN"] % 2 != 0):
        return "ERR: 2N needs an even number of units in RedGroup"
    if ne(s, "N") and eq(g, ""):
        return "ERR: redundancy scheme needs a RedGroup"
    if ne(g, "") and countifs(rg, g, sch, s) != o["GrpN"]:
        return "ERR: mixed schemes in one RedGroup"
    if left(o["RatingStatus"], 4) == "FAIL":
        return o["RatingStatus"]
    if o["LoadDer_pct"] > num(dbref(db["MaxLoad"])) * 100 + EPS:
        return "FAIL: loading exceeds permissible"
    if ne(s, "N") and isnumber(o["ContLoad_pct"]) and N(o["ContLoad_pct"]) > num(dbref(db["MaxLoadCont"])) * 100 + EPS:
        return "FAIL: contingency loading exceeds permissible"
    refT, derK, refA, derA = _derate_parts(typ, db)
    if gt(N(db["Tamb"]), refT) and not isnumber(derK):
        return "REVIEW: ambient above rating basis - derating coefficient not entered"
    if gt(N(db["Alt"]), refA) and not isnumber(derA):
        return "REVIEW: altitude above rating basis - derating not entered"
    if isnumber(db["NLThr"]) and gt(o["NL_share"], db["NLThr"]) and not isnumber(t.get("Kharm_in")):
        return "REVIEW: non-linear share above threshold - harmonic derating not entered"
    if not isnumber(t.get("Z_pct")) or not isnumber(t.get("XR")):
        return "REVIEW: impedance / X/R not entered"
    if eq(t.get("Cooling"), ""):
        return "WARN: cooling not stated"
    if not isnumber(t.get("Selected_kVA")):
        return "REVIEW: rating not selected (provisional)"
    return "OK"


TX_CALC_COLUMNS = (["Check", "HVkV_e", "VLL_e", "VLN_e", "Scen_e"] + [f"S_{s}" for s in SC] +
                   ["Conn_kVA", "Conn_kW", "P_run", "Q_run", "PF_run", "Sin_run", "Design_kVA", "Gov", "Kamb", "Kalt", "Kharm", "Kderate",
                    "DesignPlus", "Req_own", "Scheme_e", "GrpN", "GrpDesign", "GrpPlus", "UnitsAvail", "Req_cont", "Req_min", "Rec_std",
                    "Eval_kVA", "DerateOK", "RatingStatus", "Cap_derated", "Load_pct", "LoadDer_pct", "Spare_kVA", "Spare_pct",
                    "ContLoad_pct"] + [f"U_{s}" for s in SC] +
                   ["I_HV", "I_LV", "I_LV_design", "ur", "ux", "Reg_pct", "Ssc_tx", "Ssc_bus", "Isc_kA", "MaxStart_kVA", "Dip_pct",
                    "NL_share", "Inrush_A", "L1_kVA", "L2_kVA", "L3_kVA", "Unalloc_kVA", "Imb_pct", "BalCheck"])


def calculate_transformers(tx_rows, volts, loads, db, lst):
    out = [_one_tx(t, v, loads, db) for t, v in zip(tx_rows, volts)]
    _group_pass(tx_rows, out, db, lst)
    return out


def central_comparison(site_peak_S, tx_out, tx_rows, db, lst, central_type):
    a = site_peak_S
    b = sum(o["Design_kVA"] for o in tx_out if isnumber(o["Design_kVA"]))
    c = b / a if a > 0 else ""
    e = k_ambient(central_type, db) * k_altitude(central_type, db)
    f = a * (1 + N(db["Margin"])) * (1 + N(db["Future"])) / e if e > 0 else 0
    g = recommended_standard(f, lst, with_count_check=False)
    return dict(CMP_SiteDesign=a, CMP_SumTX=b, CMP_Diversity=c, CMP_Type=central_type, CMP_K=e, CMP_CentMin=f,
                CMP_CentRec=g, CMP_CentN1=f, CMP_DistInst=sum(o["Eval_kVA"] for o in tx_out if isnumber(o["Eval_kVA"])),
                CMP_NTX=countif([t.get("TX_ID") for t in tx_rows], "?*"))
