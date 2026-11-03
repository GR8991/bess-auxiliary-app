import streamlit as st

from ui import state, tables
from ui.theme import kpi, banner
from visualizations.dashboard import status_bar

p, res, user = state.page_setup("Engineering Validation", "Validation register V01–V28 (workbook logic preserved) and row-level findings.")
app = res["application"]
reg = app["register"]
counts = {s: sum(1 for r in reg if r["status"] == s) for s in ("PASS", "FAIL", "REVIEW", "INCOMPLETE", "NOT VERIFIED", "NOT APPLICABLE")}
c = st.columns(6)
for i, (k, v) in enumerate(counts.items()):
    c[i].markdown(kpi(k, v, "", "register checks"), unsafe_allow_html=True)
st.markdown(banner("A PASS means the workbook check is satisfied with the data entered — it is <b>not</b> proof of engineering compliance "
                   "or approval. Standards / AHJ items show NOT VERIFIED until verified against licensed documents (F-07).", "warn"),
            unsafe_allow_html=True)
t1, t2, t3 = st.tabs(["Validation register", "Row findings", "Status overview"])
with t1:
    a, b, cc = st.columns(3)
    stt = a.multiselect("Status", list(counts), default=[k for k in counts if k != "PASS"])
    sev = b.multiselect("Severity", ["High", "Medium", "Low", "-"], default=["High", "Medium", "Low", "-"])
    cats = sorted({r["category"] for r in reg})
    cat = cc.multiselect("Category", cats, default=cats)
    rows = [r for r in reg if r["status"] in stt and r["severity"] in sev and r["category"] in cat]
    from calculations import reference_data as R
    meta = {v["id"]: v for v in R.VALIDATION_ITEMS}
    view = [dict(ID=r["id"], Requirement=meta[r["id"]]["requirement"], Category=r["category"], Expected=meta[r["id"]]["criterion"],
                 Check=meta[r["id"]]["check"], Workbook_result=r["source_status"], Status=r["status"], Severity=r["severity"],
                 Recommended_action=r["action"], Reference=meta[r["id"]]["standard"], Evidence=meta[r["id"]]["evidence"], Note=r["note"])
            for r in rows]
    st.dataframe(tables.result_df(view, list(view[0]) if view else ["ID"]), use_container_width=True, hide_index=True, height=520)
    st.caption("Workbook_result = original v1.1 outcome (source_status, kept for traceability). Status = application display status (A-1). "
               "V25 is shown NOT APPLICABLE because the application has no display-slot limits (A-8); its workbook logic is retained.")
with t2:
    f = app["findings"]
    a, b, cc = st.columns(3)
    obj = a.multiselect("Object", ["load", "transformer", "feeder"], default=["load", "transformer", "feeder"])
    sev = b.multiselect("Severity ", ["High", "Medium", "Low"], default=["High", "Medium", "Low"])
    stt = cc.multiselect("Status ", ["FAIL", "INCOMPLETE", "REVIEW", "NOT VERIFIED"], default=["FAIL", "INCOMPLETE", "REVIEW", "NOT VERIFIED"])
    f = [x for x in f if x["object"] in obj and x["severity"] in sev and x["status"] in stt]
    st.dataframe(tables.result_df(f, ["object", "object_id", "code", "status", "severity", "description", "action", "workbook_check"]),
                 use_container_width=True, hide_index=True, height=520)
    st.markdown("<span class='label-app'>Row findings are application-level (NOT IN v1.1, A-9).</span>", unsafe_allow_html=True)
with t3:
    l, r = st.columns(2)
    l.plotly_chart(status_bar(counts, "Register checks by status"), use_container_width=True)
    r.plotly_chart(status_bar(app["status_counts"], "All statuses incl. row findings"), use_container_width=True)
