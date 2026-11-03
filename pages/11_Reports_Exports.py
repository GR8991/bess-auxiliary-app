import io

import streamlit as st

from calculations.load_calculations import CALC_COLUMNS
from calculations.transformer_sizing import TX_CALC_COLUMNS
from calculations.lv_distribution import FDR_CALC_COLUMNS
from data_models.project_schema import TABLE_INPUTS
from services import project_storage as PS, json_export as JE, excel_export as EX, pdf_report as PR
from ui import state, tables
from ui.theme import banner

p, res, user = state.page_setup("Reports and Exports", "Download the project, the engineering report and data tables.")
st.markdown(banner("Files are generated in memory for download only — nothing is stored on the server.", "info"), unsafe_allow_html=True)
base = (p["design_basis"].get("Project") or "Project").replace(" ", "_")
rev = p["metadata"].get("revision", "Rev00")
c = st.columns(3)
with c[0], st.container(border=True):
    st.markdown("**Project file (JSON)** — inputs + metadata; reload with *Upload existing project*.")
    st.download_button("Download project JSON", PS.to_json(p, res), PS.filename(p), "application/json", use_container_width=True, type="primary")
with c[1], st.container(border=True):
    st.markdown("**AI Agent handoff (JSON)** — primary handoff format.")
    st.download_button("Download handoff JSON", JE.to_bytes(JE.build_handoff(p, res, user)), f"{base}_BOP_handoff_{rev}.json",
                       "application/json", use_container_width=True)
with c[2], st.container(border=True):
    st.markdown("**PDF engineering summary** — sections A–G.")
    if st.button("Generate PDF", use_container_width=True):
        st.session_state["_pdf"] = PR.build_pdf(p, res, user)
    if st.session_state.get("_pdf"):
        st.download_button("Download PDF", st.session_state["_pdf"], f"{base}_Engineering_Summary_{rev}.pdf", "application/pdf",
                           use_container_width=True)
c = st.columns(3)
with c[0], st.container(border=True):
    st.markdown("**Excel — inputs and results** (static values from the engine).")
    if st.button("Generate Excel results", use_container_width=True):
        st.session_state["_xlsx"] = EX.results_workbook(p, res)
    if st.session_state.get("_xlsx"):
        st.download_button("Download Excel results", st.session_state["_xlsx"], f"{base}_Results_{rev}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
with c[1], st.container(border=True):
    st.markdown("**Excel — v1.1 formula workbook** with this project's inputs (independent re-check in Excel).")
    if st.button("Generate v1.1 workbook", use_container_width=True):
        st.session_state["_v11"] = EX.v11_workbook(p)
    if st.session_state.get("_v11"):
        st.download_button("Download v1.1 workbook", st.session_state["_v11"], f"{base}_v1.1_workbook_{rev}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    st.caption("Open in Excel and recalculate (Ctrl+Alt+F9). Scenario names/definitions use the v1.1 template text.")
with c[2], st.container(border=True):
    st.markdown("**CSV tables**")
    wb = res["workbook"]
    choices = {
        "Load schedule (inputs + results)": lambda: [dict({k: row.get(k) for k in TABLE_INPUTS["loads"]}, **{k: r[k] for k in CALC_COLUMNS})
                                                     for row, r in zip(p["loads"], wb["loads"])],
        "Transformer sizing": lambda: [dict({k: t.get(k) for k in TABLE_INPUTS["transformers"]}, **{k: o[k] for k in TX_CALC_COLUMNS})
                                       for t, o in zip(p["transformers"], wb["transformers"])],
        "LV feeders": lambda: [dict({k: f.get(k) for k in TABLE_INPUTS["feeders"]}, **{k: o[k] for k in FDR_CALC_COLUMNS})
                               for f, o in zip(p["feeders"], wb["feeders"])],
        "Scenario results": lambda: wb["scenarios"]["rows"],
        "Validation register": lambda: res["application"]["register"],
        "Row findings": lambda: res["application"]["findings"],
        "Category breakdown": lambda: wb["summary"]["categories_all"],
    }
    sel = st.selectbox("Table", list(choices))
    rows = choices[sel]()
    df = tables.result_df(rows, list(rows[0]) if rows else [])
    st.download_button("Download CSV", df.to_csv(index=False).encode("utf-8-sig"), f"{base}_{sel.split(' ')[0]}_{rev}.csv", "text/csv",
                       use_container_width=True)
