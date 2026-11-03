"""Engineering completion summary (brief §9 A–G), built ONLY from project inputs and engine results.
Used by the Summary page, the PDF report and the AI handoff."""
from __future__ import annotations

from calculations.excel_semantics import isnumber, text
from calculations.constants import SC

COMPLETED, PARTIAL, MISSING, REVIEW = "Completed", "Partially completed", "Missing mandatory inputs", "Requires engineering review"


def _real(rows):
    return [r for r in rows if not r.get("_padding")]


def completion(project, res):
    wb, app = res["workbook"], res["application"]
    out = []
    dbc = wb["design_basis_checks"].values()
    n_err = sum(1 for c in dbc if text(c).startswith("ERR"))
    n_rev = sum(1 for c in dbc if text(c).startswith("REVIEW"))
    out.append(("Design Basis", MISSING if n_err else (REVIEW if n_rev else COMPLETED),
                f"{n_err} error(s), {n_rev} review item(s) in input checks"))
    lib = _real(project.get("library") or [])
    used = {text(l.get("LibID")) for l in project.get("loads") or [] if l.get("LibID")}
    ls = wb["library_status"]
    bad = sum(1 for s in ls if text(s).startswith(("ERR", "WARN")))
    unrated_used = sum(1 for l, s in zip(lib, ls) if text(l.get("LibID")) in used and text(s).startswith("INFO"))
    out.append(("Equipment Library", MISSING if not lib else (REVIEW if bad else (PARTIAL if unrated_used else COMPLETED)),
                f"{len(lib)} items; {bad} with errors/warnings; {unrated_used} used items without a library rating (rating taken from load rows)"))
    loads = [(row, r) for row, r in zip(res["inputs"]["loads"], wb["loads"]) if not row.get("_padding")]
    act = [r for _, r in loads if r["Active"] == 1]
    errs = sum(1 for r in act if text(r["Check"]).startswith("ERR"))
    warns = sum(1 for r in act if text(r["Check"]).startswith("WARN"))
    infos = sum(1 for r in act if text(r["Check"]).startswith("INFO"))
    out.append(("Load Schedule", MISSING if (not act or errs) else (REVIEW if warns else (PARTIAL if infos else COMPLETED)),
                f"{len(act)} active rows: {errs} ERR, {warns} WARN, {infos} INFO"))
    v = wb["validation"]
    out.append(("Scenario Analysis", REVIEW if (v["V06"] != "PASS" or v["V07"] != "PASS" or v["V08"] != "PASS") else COMPLETED,
                f"Exclusivity {v['V06']}; profiles reviewed {v['V07']}; MAXSIM envelope {text(v['V08'])[:6]}"))
    txs = [(t, o) for t, o in zip(res["inputs"]["transformers"], wb["transformers"]) if text(t.get("TX_ID"))]
    terr = sum(1 for _, o in txs if text(o["Check"]).startswith("ERR"))
    tfail = sum(1 for _, o in txs if text(o["RatingStatus"]).startswith("FAIL") or text(o["Check"]).startswith("FAIL"))
    tprov = sum(1 for _, o in txs if not text(o["RatingStatus"]).startswith("SELECTED - adequate"))
    out.append(("Transformer Sizing", MISSING if (not txs or terr) else (REVIEW if (tfail or tprov) else COMPLETED),
                f"{len(txs)} transformer(s): {terr} ERR, {tfail} FAIL, {tprov} not confirmed as adequate"))
    fds = [(f, o) for f, o in zip(res["inputs"]["feeders"], wb["feeders"]) if text(f.get("FeederID"))]
    ffail = sum(1 for _, o in fds if text(o["Overall"]).startswith(("ERR", "FAIL")))
    frev = sum(1 for _, o in fds if text(o["Overall"]).startswith("REVIEW"))
    out.append(("LV Distribution", PARTIAL if not fds else (REVIEW if ffail else (PARTIAL if frev else COMPLETED)),
                f"{len(fds)} feeder(s): {ffail} ERR/FAIL, {frev} with incomplete data (screening checks)"))
    c = app["status_counts"]
    out.append(("Validation", REVIEW if c["FAIL"] else (PARTIAL if (c["INCOMPLETE"] or c["REVIEW"]) else COMPLETED),
                f"FAIL {c['FAIL']}, INCOMPLETE {c['INCOMPLETE']}, REVIEW {c['REVIEW']}, NOT VERIFIED {c['NOT VERIFIED']}"))
    return [dict(module=m, status=s, detail=d) for m, s, d in out]


def electrical(project, res):
    wb = res["workbook"]
    k = wb["summary"]["kpi"]
    loads = [r for row, r in zip(res["inputs"]["loads"], wb["loads"]) if not row.get("_padding")]
    rows = {r["Code"]: r for r in wb["scenarios"]["rows"]}
    peak = max(wb["scenarios"]["rows"], key=lambda r: r["S"])
    txs = [(t, o) for t, o in zip(res["inputs"]["transformers"], wb["transformers"]) if text(t.get("TX_ID"))]
    fds = [(f, o) for f, o in zip(res["inputs"]["feeders"], wb["feeders"]) if text(f.get("FeederID"))]
    lead_kvar = sum(r["P_" + peak["Code"]] * r["TanPhi"] for r in loads if r["TanPhi"] < 0) if peak["Code"] in SC else 0
    lag_kvar = sum(r["P_" + peak["Code"]] * r["TanPhi"] for r in loads if r["TanPhi"] > 0) if peak["Code"] in SC else 0
    imb = [o["Imb_pct"] for _, o in txs]
    fl = [o["Load_pct"] for _, o in fds if isnumber(o["Load_pct"])]
    return dict(
        connected_kW=k["SUM_ConnkW"], connected_kVA=k["SUM_ConnkVA"], operating_kW_peak=k["SUM_PeakkW"],
        reactive_kvar_peak=peak["Q"], apparent_kVA_peak=k["SUM_PeakkVA"], peak_pf=k["SUM_PeakPF"], max_demand_scenario=k["SUM_PeakName"],
        transient_kVA=k["SUM_Trans"], daily_kWh=k["SUM_DaykWh"], annual_MWh=k["SUM_YrMWh"],
        max_tx_demand_kVA=max([o["Design_kVA"] for _, o in txs] or [0]), n_transformers=len(txs), n_feeders=len(fds),
        max_tx_loading_pct=max([o["Load_pct"] for _, o in txs] or [0]), max_feeder_loading_pct=max(fl) if fl else None,
        max_phase_imbalance_pct=max(imb) if imb else None, lagging_kvar_peak=lag_kvar, leading_kvar_peak=lead_kvar,
        dc_kW=k["SUM_DCkW"], scenarios={c: dict(kW=r["P_counted"], kvar=r["Q"], kVA=r["S"], PF=r["PF"]) for c, r in rows.items()})


def transformer_table(res):
    out = []
    for t, o in zip(res["inputs"]["transformers"], res["workbook"]["transformers"]):
        if not text(t.get("TX_ID")):
            continue
        out.append(dict(TX_ID=t["TX_ID"], Design_kVA=o["Design_kVA"], Governing=o["Gov"], Req_min_kVA=o["Req_min"],
                        Recommended=o["Rec_std"], Selected=t.get("Selected_kVA"), Loading_pct=o["Load_pct"],
                        Derating="Complete" if o["DerateOK"] == 1 else "INCOMPLETE (F-04)", Result=o["RatingStatus"], Check=o["Check"]))
    return out


STUDIES = [
    ("Short-circuit study", "Prospective LV fault level in the workbook is a simplified c·S/(√3·U) screening value; a full IEC 60909 / IEEE study is required for equipment ratings."),
    ("Load-flow study", "Voltage regulation uses an approximate formula with the aggregate running PF (R-1); a load-flow study is required for the final design."),
    ("Protection coordination", "Breaker settings, inrush and selectivity are not designed by this tool."),
    ("Voltage-drop verification", "Feeder voltage drop is a screening check with user-entered cable data; verify with the final cable schedule."),
]


def unresolved(project, res):
    wb, app = res["workbook"], res["application"]
    items = []
    f = app["findings"]

    def add(cat, text_, ref, sev="Medium"):
        items.append(dict(category=cat, item=text_, reference=ref, severity=sev))
    n = sum(1 for x in f if x["code"] == "L26")
    if n:
        add("Missing equipment datasheets", f"{n} load row(s) without a datasheet reference", "V05 / findings L26", "Low")
    n = sum(1 for x in f if x["code"] == "L05")
    if n:
        add("Missing equipment ratings", f"{n} load row(s) without a rating", "findings L05", "High")
    for o, t in zip(wb["transformers"], res["inputs"]["transformers"]):
        if text(t.get("TX_ID")) and o["DerateOK"] == 0:
            add("Missing derating coefficients", f"{t['TX_ID']}: derating data incomplete - sizing not confirmed (F-04)", "tblTX DerateOK / V11, V13", "High")
    for key in ("DerK_L", "DerAlt_L", "DerK_D", "DerAlt_D"):
        if not isnumber(wb["design_basis"][key]):
            add("Missing derating coefficients", f"Design Basis DB_{key} not entered", "Design Basis", "Low")
    if wb["validation"]["V15"] != "PASS":
        add("Missing harmonic information", "Non-linear threshold or harmonic derating (IEEE C57.110 / manufacturer) not resolved", "V15")
    if wb["validation"]["V24"] == "REVIEW" or wb["validation"]["V17"] != "PASS":
        add("Incomplete motor-starting data", "Starting values are screening-level only; detailed motor-starting / restoration study required (F-05)", "V17, V24")
    if wb["validation"]["V18"] != "PASS":
        add("Incomplete motor-starting data", "Transformer energisation inrush data not obtained for all transformers", "V18", "Low")
    if wb["validation"]["V28"] != "PASS":
        add("Unverified standards", "Standard editions and clauses not verified against licensed copies (F-07)", "V28")
    for vid, lab in (("V26", "Fire-pump supply arrangement to be verified with AHJ / code"), ("V27", "ESS safety-system supply reliability to be verified with AHJ / code")):
        if wb["validation"][vid] == "REVIEW":
            add("Unverified standards", lab, vid)
    for name, why in STUDIES:
        add("Required detailed studies", f"{name}: {why}", "Methodology / limitations", "Medium")
    shown = {"V05", "V11", "V13", "V15", "V17", "V18", "V24", "V26", "V27", "V28"}
    for r in app["register"]:
        if r["id"] not in shown and r["status"] in ("FAIL", "REVIEW", "INCOMPLETE", "NOT VERIFIED"):
            add("Other unresolved validation items", f"{r['id']} {r['category']}: {r['status']} - {r['action']}", r["id"], r["severity"] if r["severity"] != "-" else "Medium")
    if any(t["status"] != "OK" for t in res["input_trace"]):
        add("Other unresolved validation items", "One or more input-cell formulas could not be evaluated", "Input trace", "High")
    return items


def build(project, res):
    wb, app = res["workbook"], res["application"]
    db = wb["design_basis"]
    info = project.get("project_info") or {}
    meta = project.get("metadata") or {}
    return dict(
        project=dict(name=db.get("Project"), client=info.get("client"), location=db.get("Location"),
                     country_grid_region=info.get("country_grid_region"), capacity_MW=db.get("BESS_MW"), energy_MWh=db.get("BESS_MWh"),
                     duration_h=wb["design_basis_derived"]["Duration_h"], poi_voltage_kV=info.get("poi_voltage_kv"),
                     aux_voltage_V=db.get("VLL"), revision=meta.get("revision"), document_rev=db.get("Rev"), prepared_by=db.get("Author"),
                     reviewed_by=info.get("reviewed_by"), date=info.get("date"), project_id=meta.get("project_id")),
        completion=completion(project, res), electrical=electrical(project, res), transformers=transformer_table(res),
        validation=dict(total=len(app["register"]), counts=app["status_counts"],
                        register_counts={s: sum(1 for r in app["register"] if r["status"] == s) for s in
                                         ("PASS", "FAIL", "REVIEW", "INCOMPLETE", "NOT VERIFIED", "NOT APPLICABLE")},
                        high=[x for x in app["register"] if x["severity"] == "High"] + [x for x in app["findings"] if x["severity"] == "High"]),
        unresolved=unresolved(project, res), disposition=app["disposition"],
        disposition_note="A calculated result or a completed input form is not an approved electrical design. "
                         "The highest status this application can assign is READY FOR INDEPENDENT ENGINEERING REVIEW.")
