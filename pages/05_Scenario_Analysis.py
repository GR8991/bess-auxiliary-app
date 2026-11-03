import streamlit as st

from calculations.constants import SC
from data_models.project_schema import TABLE_INPUTS
from ui import state, tables
from ui.theme import kpi, fmt, banner
from visualizations.dashboard import pqs_by_scenario, tx_utilisation, SCEN_LABEL

p, res, user = state.page_setup("Scenario Analysis", "Seven operating scenarios with participation profiles (sheet 4).")
wb = res["workbook"]
sc = wb["scenarios"]
c = st.columns(4)
pk = max(sc["rows"], key=lambda r: r["S"])
c[0].markdown(kpi("Maximum-demand scenario", pk["Scenario"] or pk["Code"], "", f"{fmt(pk['S'])} kVA"), unsafe_allow_html=True)
c[1].markdown(kpi("Peak active power", fmt(pk["P_counted"]), "kW", f"PF {fmt(pk['PF'], 3)}"), unsafe_allow_html=True)
c[2].markdown(kpi("Exclusivity accounting", sc["consistency"], "", "counted ≤ raw in every scenario"), unsafe_allow_html=True)
c[3].markdown(kpi("MAXSIM envelope", str(sc["maxsim_check"])[:6], "", "MAXSIM ≥ NORM, MAXCOOL"), unsafe_allow_html=True)
if p["design_basis"].get("ProfRev") != "Yes":
    st.markdown(banner("Scenario profiles are generic engineering assumptions — review them for this project and confirm in the Design Basis "
                       "(V07).", "warn"), unsafe_allow_html=True)
t1, t2, t3 = st.tabs(["Results", "Profiles & definitions", "Transformers & feeders by scenario"])
with t1:
    st.plotly_chart(pqs_by_scenario(sc["rows"]), use_container_width=True)
    cols = ["Code", "Scenario", "P_raw", "P_counted", "P_excluded", "Q", "S", "S_arith", "PF", "StartDelta", "S_transient", "Rows",
            "ExclResolved", "Rank"]
    st.dataframe(tables.result_df(sc["rows"], cols), use_container_width=True, hide_index=True)
    st.caption("S vector = √(ΣP² + ΣQ²) with signed Q (F-06). Transient = S + largest starting increment of loads active in the scenario, "
               "site-wide, one start at a time (R-2, screening).")
with t2:
    st.markdown("**Participation profiles** (factor 0–1 per scenario)")
    tables.edit_table("profiles", p["profiles"], TABLE_INPUTS["profiles"], {}, "prof", height=420, id_col="Profile")
    st.markdown("**Scenario definitions** (codes are fixed by the methodology; names and definitions are editable)")
    tables.edit_table("scenarios", p["scenarios"], TABLE_INPUTS["scenarios"], {}, "scen", height=300)
with t3:
    tx = [o | {"TX_ID": t.get("TX_ID")} for t, o in zip(res["inputs"]["transformers"], wb["transformers"]) if t.get("TX_ID")]
    st.plotly_chart(tx_utilisation(tx, SC), use_container_width=True)
    st.dataframe(tables.result_df(tx, ["TX_ID"] + [f"S_{s}" for s in SC] + ["Design_kVA", "Gov"]), use_container_width=True, hide_index=True)
    fd = [o | {"FeederID": f.get("FeederID")} for f, o in zip(res["inputs"]["feeders"], wb["feeders"]) if f.get("FeederID")]
    st.markdown("**Feeder apparent power by scenario (kVA)**")
    st.dataframe(tables.result_df(fd, ["FeederID"] + [f"S_{s}" for s in SC] + ["S_design", "Ib_A"]), use_container_width=True, hide_index=True)
