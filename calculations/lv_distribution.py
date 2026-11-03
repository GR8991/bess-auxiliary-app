"""Sheet 7 - LV Distribution (tblFdr, 31 calculated columns).  Preliminary SCREENING checks only;
Iz, R, X and device data are user inputs.  Voltage drop uses the feeder's aggregate running PF and
signed sinφ exactly as v1.1 (R-1)."""
from __future__ import annotations

import math

from .excel_semantics import SQRT3, isnumber, N, eq, ne, left, sumifs, countif, countifs, match_exact
from .constants import SC
from .load_calculations import dbref

FDR_INPUTS = ['FeederID', 'Board', 'TX_ID', 'Description', 'Phase', 'Conn', 'Basis', 'Length_m', 'Cable', 'Iz_A', 'R_ohm_km',
              'X_ohm_km', 'In_A', 'I2_ratio', 'Icu_kA', 'Notes']
FDR_CALC_COLUMNS = (["Overall", "VLL", "VLN", "V_fdr", "Basis_e"] + [f"S_{s}" for s in SC] +
                    ["Conn_kVA", "Loads", "P_run", "Q_run", "PF_run", "Sin_run", "S_design", "Ib_A", "In_min", "Chk_In", "Chk_Iz",
                     "Chk_I2", "VD_pct", "Chk_VD", "MaxStart_kVA", "IStart_A", "Isc_bus_kA", "Chk_Icu", "Load_pct"])


def calculate_feeders(fdr_rows, loads, tx_rows, tx_out, db):
    tx_ids = [t.get("TX_ID") for t in tx_rows]
    fids = [f.get("FeederID") for f in fdr_rows]
    lf = [r["_Feeder"] for r in loads]
    out = []
    for f in fdr_rows:
        o = {}
        fid = f.get("FeederID")
        blank = eq(fid, "")
        m = match_exact(f.get("TX_ID"), tx_ids)
        o["VLL"] = tx_out[m - 1]["VLL_e"] if m else N(db["VLL"])
        o["VLN"] = tx_out[m - 1]["VLN_e"] if m else N(db["VLN"])
        o["V_fdr"] = (o["VLL"] if eq(f.get("Conn"), "L-L") else o["VLN"]) if eq(f.get("Phase"), "1PH") else o["VLL"]
        o["Basis_e"] = "ENVELOPE" if eq(f.get("Basis"), "") else f["Basis"]
        mask = [eq(x, fid) for x in lf]
        for s in SC:
            if blank:
                o["S_" + s] = 0
            else:
                P = sumifs([r["P_" + s] for r in loads], lf, fid)
                Q = sum(r["P_" + s] * r["TanPhi"] for r, mm in zip(loads, mask) if mm)
                o["S_" + s] = math.sqrt(P ** 2 + Q ** 2)
        o["Conn_kVA"] = 0 if blank else sumifs([r["Conn_kVA"] for r in loads], lf, fid)
        o["Loads"] = 0 if blank else countifs(lf, fid, [r["Valid"] for r in loads], 1)
        o["P_run"] = 0 if blank else sumifs([r["Run_kW"] for r in loads], lf, fid)
        o["Q_run"] = 0 if blank else sum(r["Run_kW"] * r["TanPhi"] for r, mm in zip(loads, mask) if mm)
        P, Q = o["P_run"], o["Q_run"]
        o["PF_run"] = P / math.sqrt(P ** 2 + Q ** 2) if P > 0 else 1
        o["Sin_run"] = Q / math.sqrt(P ** 2 + Q ** 2) if P > 0 else 0
        Ss = [o["S_" + s] for s in SC]
        if eq(o["Basis_e"], "ENVELOPE"):
            o["S_design"] = max(Ss)
        else:
            k = match_exact(o["Basis_e"], SC)
            o["S_design"] = Ss[k - 1] if k else 0
        V = o["V_fdr"]
        one = eq(f.get("Phase"), "1PH")
        o["Ib_A"] = (o["S_design"] * 1000 / V if one else o["S_design"] * 1000 / (SQRT3 * V)) if (isnumber(V) and V > 0) else 0
        o["In_min"] = o["Ib_A"] * dbref(db["ContMult"])
        In, Iz, I2 = f.get("In_A"), f.get("Iz_A"), f.get("I2_ratio")
        o["Chk_In"] = "REVIEW: In not entered" if not isnumber(In) else ("PASS" if In >= o["In_min"] else "FAIL")
        o["Chk_Iz"] = ("REVIEW: Iz not entered" if not isnumber(Iz) else "REVIEW: In not entered" if not isnumber(In)
                       else ("PASS" if Iz >= In else "FAIL"))
        o["Chk_I2"] = ("REVIEW: data missing" if (not isnumber(I2) or not isnumber(Iz) or not isnumber(In))
                       else ("PASS" if I2 * In <= dbref(db["K145"]) * Iz else "FAIL"))
        Lm, R, X = f.get("Length_m"), f.get("R_ohm_km"), f.get("X_ohm_km")
        if isnumber(Lm) and isnumber(R) and isnumber(X) and isnumber(V) and V > 0:
            k = 2 if one else SQRT3
            o["VD_pct"] = k * o["Ib_A"] * Lm / 1000 * (R * o["PF_run"] + X * o["Sin_run"]) / V * 100
        else:
            o["VD_pct"] = ""
        if not isnumber(o["VD_pct"]):
            o["Chk_VD"] = "REVIEW: cable data missing"
        elif not isnumber(db["VDLim"]):
            o["Chk_VD"] = "REVIEW: VD limit not set"
        else:
            o["Chk_VD"] = "PASS" if o["VD_pct"] <= db["VDLim"] * 100 else "FAIL"
        o["MaxStart_kVA"] = 0 if blank else max([r["Start_kVA"] if mm else 0 for r, mm in zip(loads, mask)] or [0])
        o["IStart_A"] = ((o["MaxStart_kVA"] * 1000 / V if one else o["MaxStart_kVA"] * 1000 / (SQRT3 * V))
                         if (isnumber(V) and V > 0) else 0)
        o["Isc_bus_kA"] = tx_out[m - 1]["Isc_kA"] if m else ""
        Icu = f.get("Icu_kA")
        o["Chk_Icu"] = ("REVIEW: data missing" if (not isnumber(Icu) or not isnumber(o["Isc_bus_kA"]))
                        else ("PASS" if Icu >= o["Isc_bus_kA"] else "FAIL"))
        o["Load_pct"] = o["Ib_A"] / In * 100 if N(In) > 0 else ""
        chks = [o[c] for c in ("Chk_In", "Chk_Iz", "Chk_I2", "Chk_VD", "Chk_Icu")]
        if blank:
            o["Overall"] = ""
        elif countif(fids, fid) > 1:
            o["Overall"] = "ERR: duplicate FeederID"
        elif m is None:
            o["Overall"] = "ERR: TX_ID not defined"
        elif ne(f.get("Phase"), "3PH") and ne(f.get("Phase"), "1PH"):
            o["Overall"] = "ERR: phase must be 3PH or 1PH"
        elif any(eq(c, "FAIL") for c in chks):
            o["Overall"] = "FAIL: see check columns"
        elif any(left(c, 6) == "REVIEW" for c in chks):
            o["Overall"] = "REVIEW: incomplete data"
        else:
            o["Overall"] = "PASS (preliminary)"
        out.append(o)
    return out
