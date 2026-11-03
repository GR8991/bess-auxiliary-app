"""APPLICATION-LEVEL ADDITIONS - NOT INCLUDED IN WORKBOOK v1.1.

Everything here is derived from workbook-equivalent results and never modifies them:
  * A-1/A-2  display status (PASS/FAIL/REVIEW/INCOMPLETE/NOT VERIFIED/NOT APPLICABLE) + severity
  * A-8      V25 shown NOT APPLICABLE (no display slots in the app); workbook value kept
  * A-9      every issue on a load / transformer / feeder row as a separate finding
  * A-5      feeder-level phase balance (additional screening)
  * 9G       overall engineering disposition (never above READY FOR INDEPENDENT ENGINEERING REVIEW)
"""
from __future__ import annotations

from .excel_semantics import isnumber, eq, ne, lt, le, gt, sumifs, match_exact, countifs, text
from .constants import SC

LABEL = "NOT IN v1.1 - application-level"

V_META = {
    "V01": ("Design basis", "High", "Enter the missing / invalid Design Basis inputs (column F)."),
    "V02": ("Design basis", "Medium", "Enter margins, limits, derating coefficients and c-factor, or document why not applicable."),
    "V03": ("Load schedule", "High", "Resolve every load row showing ERR."),
    "V04": ("Load schedule", "Medium", "Resolve or accept each load-row warning."),
    "V05": ("Traceability", "Low", "Add datasheet / document references for every rated load row."),
    "V06": ("Scenario", "High", "Investigate exclusivity accounting (counted kW exceeds raw kW)."),
    "V07": ("Scenario", "Low", "Review scenario participation profiles for this project and confirm in Design Basis."),
    "V08": ("Scenario", "Medium", "Check profiles: MAXSIM should envelope NORM and MAXCOOL."),
    "V09": ("Topology", "Medium", "Assign every AC load row to a transformer."),
    "V10": ("Transformer", "High", "Correct transformer definition errors."),
    "V11": ("Transformer", "High", "Reduce load, select a larger rating, or obtain derating data."),
    "V12": ("Redundancy", "High", "Increase unit rating or units so contingency loading is within the limit."),
    "V13": ("Transformer", "High", "Enter engineer-selected ratings that are not below the calculated minimum."),
    "V14": ("Derating", "Medium", "Enter ambient / altitude derating coefficients from manufacturer or standard."),
    "V15": ("Derating", "Medium", "Set the non-linear threshold and enter harmonic derating (IEEE C57.110 / manufacturer)."),
    "V16": ("Voltage", "Medium", "Enter the regulation limit and transformer impedance / X/R."),
    "V17": ("Motor starting", "Medium", "Enter the dip limit and source/impedance data; perform a motor-starting study for critical motors."),
    "V18": ("Transformer", "Low", "Obtain energisation inrush data from the manufacturer for protection settings."),
    "V19": ("Feeder", "High", "Enter / correct In and Iz for every feeder."),
    "V20": ("Feeder", "Medium", "Enter I2/In ratio and Iz; confirm I2 <= 1.45 Iz."),
    "V21": ("Feeder", "Medium", "Enter cable length, R, X and the voltage-drop limit."),
    "V22": ("Feeder", "High", "Enter device breaking capacity and confirm Icu >= prospective fault level."),
    "V23": ("Transformer", "Medium", "Allocate single-phase loads to phases and set the imbalance limit."),
    "V24": ("Motor starting", "Medium", "Workbook starting results are screening only - perform a detailed starting / restoration study."),
    "V25": ("Workbook integrity", "Low", "Workbook display-slot limit; not applicable to the application."),
    "V26": ("Code", "Medium", "Verify fire-pump supply arrangement with the AHJ / applicable code."),
    "V27": ("Code", "Medium", "Verify ESS safety-system supply reliability with the AHJ / applicable code."),
    "V28": ("Standards", "Low", "Verify standard editions and clauses against licensed copies; confirm in Design Basis."),
}
INCOMPLETE_ON_REVIEW = {"V02", "V07", "V16", "V17", "V19", "V20", "V21", "V22"}
NOT_VERIFIED_ON_REVIEW = {"V26", "V27", "V28"}
HIGH_ONLY_ON_FAIL = {"V01", "V03", "V06", "V10", "V11", "V12", "V13", "V19", "V22"}


def display_status(vid, source, wb):
    s = text(source)
    if vid == "V25":
        return "NOT APPLICABLE"
    if s.startswith("PASS"):
        return "PASS"
    if s.startswith("NOT APPLICABLE"):
        return "NOT APPLICABLE"
    if s.startswith("FAIL"):
        return "FAIL"
    if vid == "V03":
        return "INCOMPLETE"                    # REVIEW only when no valid rows
    if vid in INCOMPLETE_ON_REVIEW:
        return "INCOMPLETE"
    if vid in NOT_VERIFIED_ON_REVIEW:
        return "NOT VERIFIED"
    return "REVIEW"


def severity(vid, status):
    if status in ("PASS", "NOT APPLICABLE"):
        return "-"
    base = V_META[vid][1]
    if status == "FAIL":
        return "High" if vid in HIGH_ONLY_ON_FAIL else ("Medium" if base == "Low" else base)
    return "Medium" if base == "High" else base


# ------------------------------------------------------------------ A-9 row findings
def load_issues(row, r, ctx_tx_ids, ctx_fdr_ids, ctx_fdr_tx, cat_ids, all_rows):
    """Every applicable condition of the workbook Check chain, evaluated independently."""
    if r["Active"] == 0:
        return []
    out = []
    ph, rt, pf, eff, qi = r["Phase_e"], r["RType_e"], r["PF_e"], r["Eff_e"], row.get("QtyInst")
    add = lambda code, st, sev, msg, act: out.append(dict(code=code, status=st, severity=sev, description=msg, action=act))
    if r.get("_input_error"):
        add("L00", "FAIL", "High", "Input formula could not be evaluated", "Correct the input formula (see input trace).")
    if r["LibIdx"] == -1:
        add("L01", "FAIL", "High", "LibID not in Equipment Library", "Select an existing LibID or add it to the library.")
    if eq(ph, ""):
        add("L02", "FAIL", "High", "Phase missing", "Enter 3PH, 1PH or DC (row or library).")
    elif ne(ph, "3PH") and ne(ph, "1PH") and ne(ph, "DC"):
        add("L03", "FAIL", "High", "Invalid phase", "Use 3PH, 1PH or DC.")
    if ne(rt, "kW_in") and ne(rt, "kVA") and ne(rt, "A") and ne(rt, "kW_out"):
        add("L04", "FAIL", "High", "Rating type missing/invalid", "Select kW_in, kVA, A or kW_out.")
    if not isnumber(r["Rating_e"]):
        add("L05", "INCOMPLETE", "High", "Rating missing (row or library)", "Enter the rating from the equipment datasheet.")
    elif r["Rating_e"] < 0:
        add("L06", "FAIL", "High", "Negative rating", "Correct the rating.")
    if (not isnumber(pf)) or le(pf, 0) or gt(pf, 1):
        add("L07", "FAIL" if isnumber(pf) else "INCOMPLETE", "High", "PF must satisfy 0 < PF <= 1", "Enter the running PF from the datasheet.")
    if eq(rt, "kW_out") and ((not isnumber(eff)) or le(eff, 0) or gt(eff, 1)):
        add("L08", "INCOMPLETE" if not isnumber(eff) else "FAIL", "High", "Efficiency 0<η<=1 required for kW_out", "Enter motor/drive efficiency.")
    if (not isnumber(qi)) or qi <= 0:
        add("L09", "FAIL", "High", "Installed quantity must be > 0", "Enter installed quantity.")
    elif lt(r["QtyDuty_e"], 0) or gt(r["QtyDuty_e"], qi):
        add("L10", "FAIL", "High", "Duty qty must be 0..installed qty", "Correct duty quantity.")
    if lt(r["DF_e"], 0) or gt(r["DF_e"], 1) or lt(r["DutyF_e"], 0) or gt(r["DutyF_e"], 1):
        add("L11", "FAIL", "High", "DF / duty factor outside 0-1", "Correct the factor.")
    if lt(r["Hours_e"], 0) or gt(r["Hours_e"], 24):
        add("L12", "FAIL", "High", "Hours/day outside 0-24", "Correct operating hours.")
    if r["PfIdx"] == 0:
        add("L13", "FAIL", "High", "Scenario profile missing/invalid", "Select a profile (row or library).")
    tx, fd = row.get("TX"), row.get("Feeder")
    if ne(tx, "") and match_exact(tx, ctx_tx_ids) is None:
        add("L14", "FAIL", "High", "Transformer not defined", "Define the transformer or correct the TX ID.")
    m = match_exact(fd, ctx_fdr_ids) if ne(fd, "") else None
    if ne(fd, "") and m is None:
        add("L15", "FAIL", "High", "Feeder not defined", "Define the feeder or correct the feeder ID.")
    if m and ne(tx, "") and ne(ctx_fdr_tx[m - 1], tx):
        add("L16", "FAIL", "High", "Feeder is supplied from a different transformer", "Align load TX with the feeder's TX.")
    if (not isnumber(r["V_app"])) or le(r["V_app"], 0):
        add("L17", "INCOMPLETE", "High", "Voltage undefined", "Set LV / DC voltage or a voltage override.")
    rc = row.get("Reactive")
    if ne(rc, "") and ne(rc, "Lag") and ne(rc, "Lead"):
        add("L18", "FAIL", "High", "Reactive must be Lag or Lead", "Select Lag or Lead.")
    g = row.get("ExclGroup")
    if ne(g, ""):
        grp = [x.get("ExclGroup") for x in all_rows]
        n = (countifs(grp, g, [x.get("TX") for x in all_rows], "<>" + text(tx)) +
             countifs(grp, g, [x.get("Feeder") for x in all_rows], "<>" + text(fd)))
        if n > 0:
            add("L19", "FAIL", "High", "ExclGroup members must share the same TX and Feeder (F-02)",
                "Place all members of the group on one TX and feeder, or split the group.")
    if eq(tx, "") and ne(ph, "DC"):
        add("L20", "REVIEW", "Medium", "No transformer assigned (site totals only)", "Assign a transformer.")
    if eq(r["Cat_e"], ""):
        add("L21", "REVIEW", "Low", "Category missing", "Assign a load category.")
    elif match_exact(r["Cat_e"], cat_ids) is None:
        add("L22", "REVIEW", "Low", "Category not in category list", "Add the category or correct it.")
    if eq(ph, "1PH") and eq(row.get("Conn1Ph"), ""):
        add("L23", "REVIEW", "Low", "1PH connection blank - L-N assumed", "Set L-N or L-L.")
    if eq(ph, "DC"):
        add("L24", "REVIEW", "Low", "DC load - excluded from AC totals (supply via charger row)", "Confirm the charger/UPS row covers this load.")
    if (not isnumber(row.get("DF"))) or (not isnumber(row.get("DutyF"))) or (not isnumber(row.get("Hours"))):
        add("L25", "REVIEW", "Low", "Design-basis default DF/duty/hours applied", "Enter row-specific factors.")
    if eq(row.get("Reference"), ""):
        add("L26", "INCOMPLETE", "Low", "No datasheet reference", "Add the equipment datasheet reference.")
    if isnumber(r["StartM_e"]) and r["StartM_e"] > 1:
        add("L27", "NOT VERIFIED", "Low", "Starting value is screening-level only (F-05)", "Confirm with manufacturer / starting study.")
    return out


def feeder_phase_balance(fdr_rows, loads, db):
    """A-5 ADDITIONAL SCREENING (NOT IN v1.1): per-feeder phase allocation using the v1.1
    transformer method (3PH/3 each phase; 1PH L-N to assigned phase; 1PH L-L split 50/50)."""
    lf = [r["_Feeder"] for r in loads]
    rk = [r["Run_kVA"] for r in loads]
    ph = [r["Phase_e"] for r in loads]
    pa = [r["_PhaseAssign"] for r in loads]
    out = []
    for f in fdr_rows:
        fid = f.get("FeederID")
        if eq(fid, ""):
            continue
        three = sumifs(rk, lf, fid, ph, "3PH") / 3
        one = lambda p: sumifs(rk, lf, fid, ph, "1PH", pa, p)
        L1 = three + one("L1") + (one("L1-L2") + one("L3-L1")) / 2
        L2 = three + one("L2") + (one("L1-L2") + one("L2-L3")) / 2
        L3 = three + one("L3") + (one("L2-L3") + one("L3-L1")) / 2
        un = one("")
        tot = L1 + L2 + L3
        imb = (max(L1, L2, L3) - tot / 3) / (tot / 3) * 100 if tot > 0 else 0
        if un > 0:
            st = "REVIEW: 1PH loads without phase allocation"
        elif not isnumber(db["ImbLim"]):
            st = "REVIEW: limit not set"
        else:
            st = "PASS" if imb <= db["ImbLim"] * 100 else "FAIL"
        out.append(dict(FeederID=fid, L1_kVA=L1, L2_kVA=L2, L3_kVA=L3, Unalloc_kVA=un, Imb_pct=imb, Check=st, label=LABEL))
    return out


def _tx_issues(t, o):
    c = text(o["Check"])
    if c in ("", "OK"):
        return []
    st = "FAIL" if c.startswith(("ERR", "FAIL")) else ("INCOMPLETE" if "not entered" in c or "not selected" in c else "REVIEW")
    issues = [dict(code="T-CHECK", status=st, severity="High" if st == "FAIL" else "Medium", description=c,
                   action="See transformer check; workbook reports the first issue only.")]
    rs = text(o["RatingStatus"])
    if rs.startswith("PROVISIONAL") or rs.startswith("REVIEW"):
        issues.append(dict(code="T-RATING", status="REVIEW", severity="Medium", description=rs,
                           action="Engineer to select a rating after manufacturer/client review."))
    if o["DerateOK"] == 0:
        issues.append(dict(code="T-DERATE", status="INCOMPLETE", severity="High", description="Derating data incomplete (F-04)",
                           action="Enter ambient/altitude coefficients and harmonic factor as applicable."))
    if text(o["BalCheck"]) not in ("", "PASS"):
        issues.append(dict(code="T-BAL", status="FAIL" if o["BalCheck"] == "FAIL" else "REVIEW", severity="Medium",
                           description=f"Phase balance: {o['BalCheck']}", action="Allocate 1PH loads / rebalance."))
    return issues


def build(res):
    wb = res["workbook"]
    inp = res["inputs"]
    register = []
    for vid, src in wb["validation"].items():
        ds = display_status(vid, src, wb)
        cat, _, act = V_META[vid]
        register.append(dict(id=vid, category=cat, source_status=src, status=ds, severity=severity(vid, ds),
                             action=act if ds not in ("PASS", "NOT APPLICABLE") else "",
                             note=("Workbook display-slot check; app has no slot limits (A-8). Workbook value retained."
                                   if vid == "V25" else "")))
    tx_ids = [t.get("TX_ID") for t in inp["transformers"]]
    fdr_ids = [f.get("FeederID") for f in inp["feeders"]]
    fdr_tx = [f.get("TX_ID") for f in inp["feeders"]]
    cat_ids = [c.get("Category") for c in inp["categories"]]
    findings = []
    for i, (row, r) in enumerate(zip(inp["loads"], wb["loads"])):
        for f in load_issues(row, r, tx_ids, fdr_ids, fdr_tx, cat_ids, inp["loads"]):
            findings.append(dict(f, object="load", object_id=text(row.get("ID")) or f"row {i + 1}", workbook_check=r["Check"]))
    for t, o in zip(inp["transformers"], wb["transformers"]):
        if ne(t.get("TX_ID"), ""):
            for f in _tx_issues(t, o):
                findings.append(dict(f, object="transformer", object_id=t["TX_ID"], workbook_check=o["Check"]))
    for f_, o in zip(inp["feeders"], wb["feeders"]):
        ov = text(o["Overall"])
        if ov and not ov.startswith("PASS"):
            findings.append(dict(code="F-CHECK", status="FAIL" if ov.startswith(("ERR", "FAIL")) else "INCOMPLETE",
                                 severity="High" if ov.startswith(("ERR", "FAIL")) else "Medium", description=ov,
                                 action="Complete / correct feeder data (screening checks).", object="feeder",
                                 object_id=f_.get("FeederID"), workbook_check=ov))
    fbal = feeder_phase_balance(inp["feeders"], wb["loads"], wb["design_basis"])
    statuses = [x["status"] for x in register] + [f["status"] for f in findings]
    n_valid = sum(r["Valid"] for r in wb["loads"])
    n_tx = sum(1 for t in tx_ids if t not in (None, ""))
    if wb["design_basis_derived"]["DB_ErrCount"] > 0 or n_valid == 0 or n_tx == 0:
        disp = "INPUTS INCOMPLETE"
    elif "FAIL" in statuses:
        disp = "CALCULATION COMPLETED — REVIEW REQUIRED"
    elif "INCOMPLETE" in statuses or "REVIEW" in statuses:
        disp = "PRELIMINARY ENGINEERING REVIEW"
    else:
        disp = "READY FOR INDEPENDENT ENGINEERING REVIEW"
    return dict(label=LABEL, register=register, findings=findings, feeder_phase_balance=fbal, disposition=disp,
                status_counts={s: statuses.count(s) for s in ("PASS", "FAIL", "REVIEW", "INCOMPLETE", "NOT VERIFIED", "NOT APPLICABLE")})
