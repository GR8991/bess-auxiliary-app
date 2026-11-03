import streamlit as st

from calculations import reference_data as R
from calculations.constants import SC
from data_models.project_schema import TABLE_INPUTS
from ui import state, tables
from ui.theme import banner, chip

p, res, user = state.page_setup("Auxiliary Load Schedule", "Equipment-level load schedule (sheet 3). Blank cells inherit library values "
                                "where a LibID is selected. Quantities accept Excel-style formulas such as =DB_N_Encl/2 (A-6).")
L = R.LISTS
wb = res["workbook"]
loads_out = [r for row, r in zip(res["inputs"]["loads"], wb["loads"]) if not row.get("_padding")]
n_err = sum(1 for r in loads_out if str(r["Check"]).startswith("ERR"))
n_warn = sum(1 for r in loads_out if str(r["Check"]).startswith("WARN"))
st.markdown(banner(f"{len(p['loads'])} rows · <b>{sum(r['Valid'] for r in loads_out)}</b> valid · <b>{n_err}</b> ERR (excluded from totals, R-5) · "
                   f"<b>{n_warn}</b> WARN. Mutually exclusive loads: give them the same ExclGroup — the highest-kVA member counts per "
                   "scenario with its own P/Q (F-01); all members must share one TX and feeder (F-02).", "bad" if n_err else "info"),
            unsafe_allow_html=True)
t1, t2, t3, t4 = st.tabs(["Inputs", "Calculated results", "Row findings", "Input formulas"])
with t1:
    opts = dict(LibID=state.ids(p["library"], "LibID"), Category=state.ids(p["categories"], "Category"), TX=state.ids(p["transformers"], "TX_ID"),
                Feeder=state.ids(p["feeders"], "FeederID"), Phase=L["lstPhase"], Conn1Ph=L["lstConn"], PhaseAssign=L["lstPhAssign"],
                RatingType=L["lstRType"], Reactive=L["lstLagLead"], Profile=state.ids(p["profiles"], "Profile"), SupplyClass=L["lstSupply"],
                Nonlinear=L["lstNL"])
    tables.edit_table("loads", p["loads"], TABLE_INPUTS["loads"], opts, "loads", height=560, id_col="ID")
with t2:
    groups = {
        "Electrical per unit": ["Phase_e", "RType_e", "Rating_e", "PF_e", "V_app", "S_unit", "P_unit", "I_unit", "TanPhi"],
        "Connected / demand / energy": ["QtyDuty_e", "DF_e", "Conn_kW", "Conn_kVA", "Run_kW", "Run_kVA", "DC_kW", "Day_kWh"],
        "Starting (screening, F-05)": ["StartM_e", "Start_kVA", "Start_kW", "StartDelta"],
        "Scenario counted kW (P_s)": [f"P_{s}" for s in SC],
        "Scenario raw kW / kVA": [f"R_{s}" for s in SC] + [f"RS_{s}" for s in SC],
        "Effective values": ["Cat_e", "Prof_e", "NL_e", "Eff_e", "StartPF_e", "DutyF_e", "Hours_e", "VLL", "VLN", "LibIdx", "PfIdx", "Valid"],
    }
    g = st.radio("Column group", list(groups), horizontal=True)
    rows = [dict({"ID": row.get("ID"), "TX": row.get("TX"), "Check": r["Check"]}, **{c: r[c] for c in groups[g]})
            for row, r in zip(p["loads"], loads_out)]
    flt = st.text_input("Filter (ID / TX / check text)", "")
    if flt:
        rows = [r for r in rows if flt.lower() in " ".join(str(v) for v in r.values()).lower()]
    st.dataframe(tables.result_df(rows, ["ID", "TX", "Check"] + groups[g]), use_container_width=True, hide_index=True, height=520)
    st.caption("All values in kW / kVA / A / V. 'Check' is the workbook's first-issue result for the row.")
with t3:
    f = [x for x in res["application"]["findings"] if x["object"] == "load"]
    sev = st.multiselect("Severity", ["High", "Medium", "Low"], default=["High", "Medium", "Low"])
    stt = st.multiselect("Status", ["FAIL", "INCOMPLETE", "REVIEW", "NOT VERIFIED"], default=["FAIL", "INCOMPLETE", "REVIEW", "NOT VERIFIED"])
    f = [x for x in f if x["severity"] in sev and x["status"] in stt]
    st.markdown("<span class='label-app'>Application-level findings (NOT IN v1.1, A-9): every issue on a row is listed separately; "
                "the workbook's first-issue result is kept alongside.</span>", unsafe_allow_html=True)
    st.dataframe(tables.result_df(f, ["object_id", "code", "status", "severity", "description", "action", "workbook_check"]),
                 use_container_width=True, hide_index=True, height=480)
with t4:
    tr = res["input_trace"]
    if tr:
        st.dataframe(tables.result_df(tr, ["where", "field", "expression", "value", "references", "status"]), use_container_width=True,
                     hide_index=True)
    else:
        st.info("No input-cell formulas in this project.")
    st.caption("Input formulas are evaluated like Excel input cells (a reference to a blank Design Basis cell = 0). Unsupported formulas are "
               "reported and the row is excluded — never replaced by a number.")
