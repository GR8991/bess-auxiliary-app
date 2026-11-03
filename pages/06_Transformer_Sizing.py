import streamlit as st

from calculations import reference_data as R
from calculations.constants import SC
from data_models.project_schema import TABLE_INPUTS
from ui import state, tables
from ui.theme import banner, chip, fmt, kpi
from visualizations.dashboard import tx_loading, tx_utilisation, phase_balance

p, res, user = state.page_setup("Transformer Sizing", "Centralized or distributed auxiliary transformers, N / N+1 / 2N (sheet 6).")
wb = res["workbook"]
L = R.LISTS
st.markdown(banner("Three separate values are reported: <b>calculated minimum</b> (Req_min), <b>recommended standard rating</b> (Rec_std — "
                   "NOT an approved selection) and the <b>engineer-selected rating</b> (Selected_kVA). A rating can only show "
                   "'SELECTED - adequate' when derating data are complete (DerateOK, F-04).", "info"), unsafe_allow_html=True)
txs = [(t, o) for t, o in zip(res["inputs"]["transformers"], wb["transformers"]) if t.get("TX_ID")]
t1, t2, t3, t4, t5 = st.tabs(["Inputs", "Sizing results", "Loading & utilisation", "Phase balance", "Screening: regulation / fault / start"])
with t1:
    opts = dict(Role=L["lstRole"], Scheme=L["lstScheme"], Type=L["lstTXType"], Cooling=L["lstCoolLiq"] + [c for c in L["lstCoolDry"] if c != "Other"],
                DesignScen=L["lstDesignScen"])
    tables.edit_table("transformers", p["transformers"], TABLE_INPUTS["transformers"], opts, "tx", height=380, id_col="TX_ID")
    st.caption("Cooling options: liquid-filled ONAN…ODAF / dry-type AN, AF, AN/AF. Impedance, X/R, harmonic factor and inrush are "
               "manufacturer data — leave blank until known (flagged).")
with t2:
    for t, o in txs:
        with st.container(border=True):
            a, b, c, d, e = st.columns([1.1, 1, 1, 1, 1.6])
            a.markdown(f"**{t['TX_ID']}** · {t.get('Role') or '–'}  \n<span class='small-muted'>{t.get('Type') or '–'} {t.get('Cooling') or ''} · "
                       f"scheme {o['Scheme_e']} · design {o['Gov']}</span>", unsafe_allow_html=True)
            b.markdown(kpi("Calculated minimum", fmt(o["Req_min"]), "kVA", f"demand {fmt(o['Design_kVA'])} kVA"), unsafe_allow_html=True)
            c.markdown(kpi("Recommended standard", fmt(o["Rec_std"]), "kVA" if isinstance(o["Rec_std"], (int, float)) else "",
                           "not approved"), unsafe_allow_html=True)
            d.markdown(kpi("Engineer-selected", fmt(t.get("Selected_kVA")), "kVA" if t.get("Selected_kVA") else "", f"loading {fmt(o['Load_pct'])} %"),
                       unsafe_allow_html=True)
            e.markdown(f"{chip(o['RatingStatus'])}<br/>{chip(o['Check'])}<br/>"
                       f"{chip('PASS' if o['DerateOK'] == 1 else 'INCOMPLETE', 'Derating complete' if o['DerateOK'] == 1 else 'Derating data incomplete')}",
                       unsafe_allow_html=True)
    cols = ["Design_kVA", "Gov", "Kamb", "Kalt", "Kharm", "Kderate", "DesignPlus", "Req_own", "GrpN", "UnitsAvail", "GrpPlus", "Req_cont",
            "Req_min", "Rec_std", "Eval_kVA", "Cap_derated", "Load_pct", "LoadDer_pct", "Spare_kVA", "Spare_pct", "ContLoad_pct", "DerateOK",
            "RatingStatus", "I_HV", "I_LV", "I_LV_design", "NL_share", "Check"]
    st.dataframe(tables.result_df([o | {"TX_ID": t["TX_ID"]} for t, o in txs], ["TX_ID"] + cols), use_container_width=True, hide_index=True)
    cc = wb["central_comparison"]
    st.markdown("**Centralized vs distributed comparison (site level)**")
    x = st.columns(5)
    x[0].markdown(kpi("Site coincident demand", fmt(cc["CMP_SiteDesign"]), "kVA"), unsafe_allow_html=True)
    x[1].markdown(kpi("Σ per-TX demand", fmt(cc["CMP_SumTX"]), "kVA", f"diversity {fmt(cc['CMP_Diversity'], 2)}"), unsafe_allow_html=True)
    x[2].markdown(kpi("Centralized minimum", fmt(cc["CMP_CentMin"]), "kVA", f"{cc['CMP_Type']} · K {fmt(cc['CMP_K'], 3)}"), unsafe_allow_html=True)
    x[3].markdown(kpi("Centralized recommended", fmt(cc["CMP_CentRec"]), "", "not approved"), unsafe_allow_html=True)
    x[4].markdown(kpi("Distributed Σ ratings", fmt(cc["CMP_DistInst"]), "kVA", f"{cc['CMP_NTX']} transformer(s)"), unsafe_allow_html=True)
with t3:
    tx = [o | {"TX_ID": t["TX_ID"]} for t, o in txs]
    st.plotly_chart(tx_loading(tx), use_container_width=True)
    st.plotly_chart(tx_utilisation(tx, SC), use_container_width=True)
    st.caption("Loading % = design demand ÷ evaluated rating (margins excluded, R-6). Contingency loading assumes equal sharing within the group.")
with t4:
    rows = [dict(TX_ID=t["TX_ID"], L1_kVA=o["L1_kVA"], L2_kVA=o["L2_kVA"], L3_kVA=o["L3_kVA"], Unalloc_kVA=o["Unalloc_kVA"],
                 Imb_pct=o["Imb_pct"], BalCheck=o["BalCheck"]) for t, o in txs]
    st.plotly_chart(phase_balance(rows, "TX_ID", "Phase loading per transformer (demand kVA)"), use_container_width=True)
    st.dataframe(tables.result_df(rows, list(rows[0]) if rows else []), use_container_width=True, hide_index=True)
with t5:
    st.markdown(banner("SCREENING ONLY. Regulation uses the aggregate running PF (R-1); fault level is c·S/(√3·U) with TX + MV source only; "
                       "the start dip assumes one start at a time. Detailed short-circuit, load-flow and motor-starting studies are required.",
                       "warn"), unsafe_allow_html=True)
    cols = ["PF_run", "Sin_run", "ur", "ux", "Reg_pct", "Ssc_tx", "Ssc_bus", "Isc_kA", "MaxStart_kVA", "Dip_pct", "Inrush_A"]
    st.dataframe(tables.result_df([o | {"TX_ID": t["TX_ID"]} for t, o in txs], ["TX_ID"] + cols), use_container_width=True, hide_index=True)
