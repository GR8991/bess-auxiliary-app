"""Sheet 8 - Validation register V01-V28 (workbook result = `source_status`, reproduced exactly),
standards-matrix applicability, and register counts.  Application-level status mapping (A-1/A-2/A-8)
lives in calculations/app_layer.py and never alters `source_status`."""
from __future__ import annotations

from .excel_semantics import isnumber, eq, countif, countifs, sumifs, num_to_text
from .load_calculations import dbref
from . import reference_data as R

EPS = 0.000001


def register(db, dbx, loads, scen, tx_rows, tx_out, fdr_rows, fdr_out, cats):
    L = loads
    chk = [r["Check"] for r in L]
    txid = [t.get("TX_ID") for t in tx_rows]
    txchk = [o["Check"] for o in tx_out]
    n_tx = countif(txid, "?*")
    fid = [f.get("FeederID") for f in fdr_rows]
    n_f = countif(fid, "?*")
    fo = lambda c: [o[c] for o in fdr_out]
    tc = lambda c: [o[c] for o in tx_out]
    v = {}
    v["V01"] = "PASS" if dbx["DB_ErrCount"] == 0 else "FAIL"
    v["V02"] = "PASS" if dbx["DB_RevCount"] == 0 else "REVIEW"
    v["V03"] = ("PASS" if sum(r["Valid"] for r in L) > 0 else "REVIEW") if countif(chk, "ERR*") == 0 else "FAIL"
    v["V04"] = "PASS" if countif(chk, "WARN*") == 0 else "REVIEW"
    v["V05"] = "PASS" if countifs([r["Valid"] for r in L], 1, [r["_Reference"] for r in L], "") == 0 else "REVIEW"
    v["V06"] = scen["consistency"]
    v["V07"] = "PASS" if eq(db["ProfRev"], "Yes") else "REVIEW"
    v["V08"] = scen["maxsim_check"]
    v["V09"] = "PASS" if countifs([r["IsAC"] for r in L], 1, [r["_TX"] for r in L], "") == 0 else "REVIEW"
    v["V10"] = "PASS" if countif(txchk, "ERR*") == 0 else "FAIL"
    lim = num_to_text(dbref(db["MaxLoad"]) * 100 + EPS)
    if n_tx == 0:
        v["V11"] = "REVIEW"
    elif countifs(txid, "?*", tc("LoadDer_pct"), ">" + lim) > 0:
        v["V11"] = "FAIL"
    elif countifs(txid, "?*", tc("DerateOK"), 0) > 0:
        v["V11"] = "REVIEW"
    else:
        v["V11"] = "PASS"
    limc = num_to_text(dbref(db["MaxLoadCont"]) * 100 + EPS)
    if countifs(tc("Scheme_e"), "<>N", txid, "?*") == 0:
        v["V12"] = "NOT APPLICABLE"
    else:
        v["V12"] = "PASS" if countif(tc("ContLoad_pct"), ">" + limc) == 0 else "FAIL"
    rs = tc("RatingStatus")
    if countif(rs, "FAIL*") > 0:
        v["V13"] = "FAIL"
    elif countif(rs, "SELECTED*") == n_tx:
        v["V13"] = "PASS" if n_tx > 0 else "REVIEW"
    else:
        v["V13"] = "REVIEW"
    v["V14"] = "PASS" if countif(txchk, "REVIEW: ambient*") + countif(txchk, "REVIEW: altitude*") == 0 else "REVIEW"
    v["V15"] = "REVIEW" if not isnumber(db["NLThr"]) else ("PASS" if countif(txchk, "REVIEW: non-linear*") == 0 else "REVIEW")
    if not isnumber(db["RegLim"]):
        v["V16"] = "REVIEW"
    elif countifs(txid, "?*", tc("Reg_pct"), "") > 0:
        v["V16"] = "REVIEW"
    else:
        v["V16"] = "PASS" if countif(tc("Reg_pct"), ">" + num_to_text(db["RegLim"] * 100)) == 0 else "FAIL"
    if not isnumber(db["DipLim"]):
        v["V17"] = "REVIEW"
    elif countif(tc("Dip_pct"), ">" + num_to_text(db["DipLim"] * 100)) == 0:
        v["V17"] = "REVIEW" if countifs(txid, "?*", tc("Dip_pct"), "") > 0 else "PASS"
    else:
        v["V17"] = "FAIL"
    v["V18"] = "PASS" if countifs(txid, "?*", [t.get("InrushMult") for t in tx_rows], "") == 0 else "REVIEW"
    ci, cz = fo("Chk_In"), fo("Chk_Iz")
    if n_f == 0:
        v["V19"] = "REVIEW"
    elif countif(ci, "FAIL") + countif(cz, "FAIL") > 0:
        v["V19"] = "FAIL"
    elif countif(ci, "REVIEW*") + countif(cz, "REVIEW*") > 0:
        v["V19"] = "REVIEW"
    else:
        v["V19"] = "PASS"
    c2 = fo("Chk_I2")
    if eq(db["Jurisdiction"], "US (NEC/NFPA/IEEE)"):
        v["V20"] = "NOT APPLICABLE"
    else:
        v["V20"] = "FAIL" if countif(c2, "FAIL") > 0 else ("REVIEW" if countif(c2, "REVIEW*") > 0 else "PASS")
    for vid, col in (("V21", "Chk_VD"), ("V22", "Chk_Icu")):
        c = fo(col)
        v[vid] = ("FAIL" if countif(c, "FAIL") > 0 else "REVIEW" if countif(c, "REVIEW*") > 0
                  else "REVIEW" if n_f == 0 else "PASS")
    bc = tc("BalCheck")
    v["V23"] = ("FAIL" if countif(bc, "FAIL") > 0 else "REVIEW" if countif(bc, "REVIEW*") > 0
                else "REVIEW" if n_tx == 0 else "PASS")
    v["V24"] = "NOT APPLICABLE" if countif([r["StartDelta"] for r in L], ">0") == 0 else "REVIEW"
    n_cat = countif([c.get("Category") for c in cats], "?*")
    v["V25"] = "PASS" if (n_tx <= 20 and n_cat <= 30) else "FAIL: display blocks truncated - extend slots"
    v["V26"] = "NOT APPLICABLE" if sumifs([r["Conn_kW"] for r in L], [r["Cat_e"] for r in L], "FIRE-PUMP") == 0 else "REVIEW"
    v["V27"] = "REVIEW"
    v["V28"] = "PASS" if eq(db["StdRev"], "Yes") else "REVIEW"
    vals = list(v.values())
    counts = {k: countif(vals, k + "*") for k in ("PASS", "FAIL", "REVIEW", "NOT APPLICABLE")}
    return v, counts


def standards_applicability(db):
    out = {}
    for s in R.STANDARDS_MATRIX:
        j = s["jurisdiction"]
        if eq(db["Jurisdiction"], ""):
            out[s["ref"]] = "REVIEW: set jurisdiction"
        elif j == "Both":
            out[s["ref"]] = "YES (if specified)"
        else:
            other = "IEC" if j == "IEC" else "US (NEC/NFPA/IEEE)"
            out[s["ref"]] = ("YES (if specified)" if (eq(db["Jurisdiction"], "Mixed / Other") or eq(db["Jurisdiction"], other))
                             else "NOT APPLICABLE (check)")
    return out


def library_status(lib):
    ids = [l.get("LibID") for l in lib]
    out = []
    for l in lib:
        if eq(l.get("LibID"), ""):
            out.append("")
        elif countif(ids, l.get("LibID")) > 1:
            out.append("ERR: duplicate LibID")
        elif isnumber(l.get("RatedValue")):
            out.append("WARN: rating has no source" if eq(l.get("Source"), "") else "OK - rated")
        else:
            out.append("INFO: rating not entered (enter in library or load row)")
    return out
