import streamlit as st

from calculations import reference_data as R
from data_models.project_schema import TABLE_INPUTS
from ui import state, tables
from ui.theme import banner
from visualizations.dashboard import feeder_loading, phase_balance

p, res, user = state.page_setup("LV Distribution", "Feeder allocation and preliminary checks (sheet 7).")
wb = res["workbook"]
L = R.LISTS
st.markdown(banner("<b>Screening calculations only.</b> Iz, R, X, In, I2/In and Icu are YOUR inputs from the applicable standard tables "
                   "and manufacturer data; the application does not select cables or breakers. Voltage drop uses the feeder's aggregate "
                   "running PF (R-1). Final cable/protection design requires detailed studies.", "warn"), unsafe_allow_html=True)
fds = [(f, o) for f, o in zip(res["inputs"]["feeders"], wb["feeders"]) if f.get("FeederID")]
t1, t2, t3 = st.tabs(["Inputs", "Feeder checks", "Phase balance (additional screening)"])
with t1:
    opts = dict(TX_ID=state.ids(p["transformers"], "TX_ID"), Phase=["3PH", "1PH"], Conn=L["lstConn"], Basis=L["lstDesignScen"])
    tables.edit_table("feeders", p["feeders"], TABLE_INPUTS["feeders"], opts, "fdr", height=380, id_col="FeederID")
with t2:
    st.plotly_chart(feeder_loading([o | {"FeederID": f["FeederID"]} for f, o in fds]), use_container_width=True)
    cols = ["Overall", "V_fdr", "Basis_e", "S_design", "Loads", "PF_run", "Ib_A", "In_min", "Chk_In", "Chk_Iz", "Chk_I2", "VD_pct", "Chk_VD",
            "MaxStart_kVA", "IStart_A", "Isc_bus_kA", "Chk_Icu", "Load_pct"]
    st.dataframe(tables.result_df([o | {"FeederID": f["FeederID"]} for f, o in fds], ["FeederID"] + cols), use_container_width=True, hide_index=True)
    if not fds or len(p["feeders"]) == 0:
        st.markdown(banner("No feeder defined. As in workbook v1.1, an empty feeder table is evaluated as one blank row, so V19–V22 report "
                           "REVIEW (finding E-1 / R-7, documented — no methodology change).", "info"), unsafe_allow_html=True)
with t3:
    fb = res["application"]["feeder_phase_balance"]
    st.markdown("<span class='label-app'>NOT IN v1.1 — application-level additional screening (approval A-5). Same allocation method as the "
                "v1.1 transformer balance, applied per feeder.</span>", unsafe_allow_html=True)
    st.plotly_chart(phase_balance(fb, "FeederID", "Phase loading per feeder (demand kVA)"), use_container_width=True)
    st.dataframe(tables.result_df(fb, ["FeederID", "L1_kVA", "L2_kVA", "L3_kVA", "Unalloc_kVA", "Imb_pct", "Check"]), use_container_width=True,
                 hide_index=True)
