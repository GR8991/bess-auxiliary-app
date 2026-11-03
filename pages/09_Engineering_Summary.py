import streamlit as st

from services import engineering_summary as ES
from ui import state, tables
from ui.theme import kpi, fmt, chip, banner
from visualizations.dashboard import category_donut, pqs_by_scenario, tx_loading, completion_chart, phase_balance, feeder_loading, status_bar

p, res, user = state.page_setup("Engineering Summary", "Consolidated completion summary generated from the project inputs and engine results.")
s = ES.build(p, res)
wb = res["workbook"]
st.markdown(f'<div class="disp"><div class="t">G · Overall engineering disposition</div><div class="v">{s["disposition"]}</div>'
            f'<div class="small-muted">{s["disposition_note"]}</div></div>', unsafe_allow_html=True)
st.write("")
st.subheader("A · Project information")
pi = s["project"]
c = st.columns(5)
for i, (lab, v, u) in enumerate([("Project", pi["name"], ""), ("Capacity", fmt(pi["capacity_MW"], 2), "MW"), ("Energy", fmt(pi["energy_MWh"], 2), "MWh"),
                                  ("Duration", fmt(pi["duration_h"], 2), "h"), ("POI voltage", fmt(pi["poi_voltage_kV"]), "kV")]):
    c[i].markdown(kpi(lab, v if v not in (None, "") else "–", u), unsafe_allow_html=True)
c = st.columns(5)
for i, (lab, v) in enumerate([("Aux system voltage", f"{fmt(pi['aux_voltage_V'], 0)} V"), ("Revision", pi["revision"]),
                              ("Prepared by", pi["prepared_by"]), ("Reviewed by", pi["reviewed_by"]), ("Date", pi["date"])]):
    c[i].markdown(kpi(lab, v or "–"), unsafe_allow_html=True)

st.subheader("B · Input completion")
l, r = st.columns([1, 1.2])
l.plotly_chart(completion_chart(s["completion"]), use_container_width=True)
with r:
    for it in s["completion"]:
        key = {"Completed": "PASS", "Partially completed": "REVIEW", "Missing mandatory inputs": "FAIL"}.get(it["status"], "NOT VERIFIED")
        st.markdown(f"{chip(key, it['status'])} **{it['module']}** — <span class='small-muted'>{it['detail']}</span>", unsafe_allow_html=True)

st.subheader("C · Electrical design summary")
e = s["electrical"]
c = st.columns(4)
c[0].markdown(kpi("Connected active load", fmt(e["connected_kW"]), "kW", f"{fmt(e['connected_kVA'])} kVA"), unsafe_allow_html=True)
c[1].markdown(kpi("Operating active load (peak)", fmt(e["operating_kW_peak"]), "kW", e["max_demand_scenario"] or ""), unsafe_allow_html=True)
c[2].markdown(kpi("Reactive load (peak)", fmt(e["reactive_kvar_peak"]), "kvar", f"lag {fmt(e['lagging_kvar_peak'])} / lead {fmt(e['leading_kvar_peak'])}"),
              unsafe_allow_html=True)
c[3].markdown(kpi("Apparent load (peak)", fmt(e["apparent_kVA_peak"]), "kVA", f"PF {fmt(e['peak_pf'], 3)}"), unsafe_allow_html=True)
c = st.columns(4)
c[0].markdown(kpi("Max transformer demand", fmt(e["max_tx_demand_kVA"]), "kVA", f"{e['n_transformers']} transformer(s)"), unsafe_allow_html=True)
c[1].markdown(kpi("Max transformer loading", fmt(e["max_tx_loading_pct"]), "%", "of evaluated rating"), unsafe_allow_html=True)
c[2].markdown(kpi("Max feeder loading", fmt(e["max_feeder_loading_pct"]), "%", f"{e['n_feeders']} feeder(s)"), unsafe_allow_html=True)
c[3].markdown(kpi("Max phase imbalance", fmt(e["max_phase_imbalance_pct"]), "%", "per transformer"), unsafe_allow_html=True)
l, r = st.columns(2)
l.plotly_chart(category_donut(wb["summary"]["categories_all"]), use_container_width=True)
r.plotly_chart(pqs_by_scenario(wb["scenarios"]["rows"]), use_container_width=True)

st.subheader("D · Transformer summary")
st.dataframe(tables.result_df(s["transformers"], ["TX_ID", "Design_kVA", "Req_min_kVA", "Recommended", "Selected", "Loading_pct", "Derating", "Result"]),
             use_container_width=True, hide_index=True,
             column_config={"Design_kVA": "Calculated kVA", "Req_min_kVA": "Calculated minimum kVA", "Recommended": "Recommended rating",
                            "Selected": "Selected rating", "Loading_pct": "Loading %", "Derating": "Derating status"})
txs = [o | {"TX_ID": t["TX_ID"]} for t, o in zip(res["inputs"]["transformers"], wb["transformers"]) if t.get("TX_ID")]
l, r = st.columns(2)
l.plotly_chart(tx_loading(txs), use_container_width=True)
r.plotly_chart(phase_balance([dict(TX_ID=t["TX_ID"], L1_kVA=t["L1_kVA"], L2_kVA=t["L2_kVA"], L3_kVA=t["L3_kVA"]) for t in txs], "TX_ID",
                             "Phase balance per transformer (kVA)"), use_container_width=True)
st.plotly_chart(feeder_loading([o | {"FeederID": f["FeederID"]} for f, o in zip(res["inputs"]["feeders"], wb["feeders"]) if f.get("FeederID")]),
                use_container_width=True)

st.subheader("E · Validation summary")
v = s["validation"]
c = st.columns(6)
c[0].markdown(kpi("Total checks", v["total"]), unsafe_allow_html=True)
for i, k in enumerate(["PASS", "FAIL", "REVIEW", "INCOMPLETE", "NOT VERIFIED"]):
    c[i + 1].markdown(kpi(k, v["register_counts"][k]), unsafe_allow_html=True)
l, r = st.columns([1, 1.4])
l.plotly_chart(status_bar(v["register_counts"]), use_container_width=True)
with r:
    hi = [x for x in res["application"]["register"] if x["severity"] == "High"]
    hf = [x for x in res["application"]["findings"] if x["severity"] == "High"]
    st.markdown(f"**Critical / high-severity findings:** {len(hi)} register item(s), {len(hf)} row finding(s)")
    act = [x for x in res["application"]["register"] if x["status"] not in ("PASS", "NOT APPLICABLE")]
    st.dataframe(tables.result_df(act, ["id", "category", "status", "severity", "action"]), use_container_width=True, hide_index=True, height=300)

st.subheader("F · Engineering limitations and unresolved items")
st.dataframe(tables.result_df(s["unresolved"], ["category", "item", "reference", "severity"]), use_container_width=True, hide_index=True)
st.markdown(banner("Unresolved items are listed explicitly — none are hidden in remarks. This summary does not approve the design.", "info"),
            unsafe_allow_html=True)
