"""Sheet 4 - Scenario Analysis results block and Sheet 5 - Summary (KPIs, scenario comparison,
category slots, TX allocation slots).  R-2: transient is site-wide (S + largest start Δ of loads
active in the scenario), added arithmetically."""
from __future__ import annotations

import math

from .excel_semantics import isnumber, N, eq, text, sumifs, countif, countifs, match_exact, rank_desc
from .constants import SC, SUMMARY_CAT_SLOTS, SUMMARY_TX_SLOTS

EPS = 0.000001


def site_results(loads, scen_rows):
    res = []
    for i, s in enumerate(SC):
        P = [r["P_" + s] for r in loads]
        tan = [r["TanPhi"] for r in loads]
        raw = sum(r["R_" + s] for r in loads)
        Psum = sum(P)
        Q = sum(p * t for p, t in zip(P, tan))
        S = math.sqrt(Psum ** 2 + Q ** 2)
        res.append(dict(
            Code=text(scen_rows[i].get("Code")) if i < len(scen_rows) else "",
            Scenario=text(scen_rows[i].get("ScenarioName")) if i < len(scen_rows) else "",
            P_raw=raw, P_counted=Psum, P_excluded=raw - Psum, Q=Q, S=S,
            S_arith=sum(p * math.sqrt(1 + t ** 2) for p, t in zip(P, tan)),
            PF=Psum / S if S > 0 else "",
            StartDelta=max([r["StartDelta"] if r["P_" + s] > 0 else 0 for r in loads] or [0]),
            Rows=countif(P, ">0"),
            ExclResolved=sum(1 for r in loads if (r["_ExclGroup"] not in (None, "")) and r["R_" + s] > 0 and r["P_" + s] == 0)))
        res[-1]["S_transient"] = res[-1]["S"] + res[-1]["StartDelta"]
    Ss = [r["S"] for r in res]
    for r in res:
        r["Rank"] = rank_desc(r["S"], Ss)
    peak = max(Ss)
    k = Ss.index(peak)
    consistency = "PASS" if sum(1 for r in res if r["P_counted"] > r["P_raw"] + EPS) == 0 else "FAIL"
    maxsim = ("PASS" if (res[2]["S"] >= res[0]["S"] - EPS and res[2]["S"] >= res[1]["S"] - EPS)
              else "REVIEW: MAXSIM lower than NORM/MAXCOOL - check profiles")
    return dict(rows=res, peak_name=res[k]["Scenario"], peak_S=peak, consistency=consistency, maxsim_check=maxsim)


def summary(loads, scen, db, cats, tx_rows, tx_out, settings):
    days = settings.get("days_per_year", 365)
    brk = settings.get("breakdown_scenario", "MAXSIM")
    Ss = [r["S"] for r in scen["rows"]]
    Ps = [r["P_counted"] for r in scen["rows"]]
    k = Ss.index(max(Ss))
    connkw = sum(r["Conn_kW"] for r in loads)
    peak_kw = Ps[k]
    peak_kva = max(Ss)
    day = sum(r["Day_kWh"] for r in loads)
    runkva = sum(r["Run_kVA"] for r in loads)
    kp = dict(
        SUM_ConnkW=connkw,
        SUM_ConnkVA=sum(r["Conn_kVA"] for r in loads),
        SUM_ConnkVAv=math.sqrt(connkw ** 2 + sum(r["Conn_kW"] * r["TanPhi"] for r in loads) ** 2),
        SUM_RunkW=sum(r["Run_kW"] for r in loads),
        SUM_PeakkVA=peak_kva, SUM_PeakkW=peak_kw,
        SUM_PeakPF=peak_kw / peak_kva if peak_kva > 0 else "",
        SUM_PeakName=scen["rows"][k]["Scenario"],
        SUM_Trans=max(r["S_transient"] for r in scen["rows"]),
        SUM_DaykWh=day, SUM_Days=days,
        SUM_YrMWh=day * N(days) / 1000,
        SUM_kWhPerMWh=day / db["BESS_MWh"] if N(db["BESS_MWh"]) > 0 else "",
        SUM_PctMW=peak_kw / (db["BESS_MW"] * 1000) * 100 if N(db["BESS_MW"]) > 0 else "",
        SUM_DCkW=sum(r["DC_kW"] for r in loads),
        SUM_Rows=f"{_g(sum(r['Valid'] for r in loads))} / {_g(sum(r['Active'] for r in loads))}",
        SUM_ErrRows=countif([r["Check"] for r in loads], "ERR*"),
        SUM_NL=sumifs([r["Run_kVA"] for r in loads], [r["NL_e"] for r in loads], "Y") / runkva if runkva > 0 else 0,
    )
    # category breakdown: workbook has 30 slots (equivalence); 'all' is the unlimited application list
    cat_ids = [c.get("Category") for c in cats]
    ce = [r["Cat_e"] for r in loads]
    m = match_exact(brk, SC)
    scen_p = Ps[m - 1] if m else None

    def cat_row(c):
        d = sumifs([r["P_" + SC[m - 1]] for r in loads], ce, c) if m else 0
        return dict(Category=c, Conn_kW=sumifs([r["Conn_kW"] for r in loads], ce, c),
                    Conn_kVA=sumifs([r["Conn_kVA"] for r in loads], ce, c), Scen_kW=d,
                    Share=(d / scen_p if (scen_p or 0) > 0 else 0),
                    Day_kWh=sumifs([r["Day_kWh"] for r in loads], ce, c),
                    Rows=countifs(ce, c, [r["Valid"] for r in loads], 1))
    slots = []
    for i in range(SUMMARY_CAT_SLOTS):
        c = cat_ids[i] if i < len(cat_ids) else None
        slots.append(cat_row(c) if (c is not None and c != "") else dict(Category=""))
    residual = dict(Conn_kW=connkw - sum(s.get("Conn_kW", 0) for s in slots),
                    Conn_kVA=kp["SUM_ConnkVA"] - sum(s.get("Conn_kVA", 0) for s in slots),
                    Scen_kW=(scen_p if m else 0) - sum(s.get("Scen_kW", 0) for s in slots),
                    Day_kWh=day - sum(s.get("Day_kWh", 0) for s in slots))
    all_cats = [cat_row(c) for c in cat_ids if c not in (None, "")]
    txs = []
    for i in range(SUMMARY_TX_SLOTS):
        t = tx_rows[i].get("TX_ID") if i < len(tx_rows) else None
        if t in (None, ""):
            txs.append(dict(TX_ID=""))
        else:
            o = tx_out[i]
            txs.append(dict(TX_ID=t, Role=text(tx_rows[i].get("Role")), Conn_kVA=o["Conn_kVA"], Design_kVA=o["Design_kVA"],
                            Gov=text(o["Gov"]), Eval_kVA=o["Eval_kVA"], Load_pct=o["Load_pct"], RatingStatus=text(o["RatingStatus"])))
    mx = max(Ss)
    comparison = [dict(Scenario=r["Scenario"], kW=r["P_counted"], kVA=r["S"], PF=(r["P_counted"] / r["S"] if r["S"] > 0 else ""),
                       PctOfPeak=(r["S"] / mx if mx > 0 else 0), Transient=r["S_transient"]) for r in scen["rows"]]
    ntx = countif([t.get("TX_ID") for t in tx_rows], "?*")
    tx_warn = ("WARNING: more than 20 transformers - list truncated (calculations unaffected, see sheet 6)" if ntx > 20 else "")
    return dict(kpi=kp, scenario_comparison=comparison, tx_truncation_warning=tx_warn, category_slots=slots, category_residual=residual, categories_all=all_cats, tx_slots=txs,
                breakdown_scenario=brk, cat_slot_truncated=sum(1 for c in cat_ids if c not in (None, "")) > SUMMARY_CAT_SLOTS)


def _g(x):
    return str(int(x)) if float(x).is_integer() else str(x)
