import streamlit as st

from data_models.project_schema import new_project, PROJECT_INFO_FIELDS, OVERVIEW_DB_LINKS
from services import project_storage as PS
from services.excel_import import import_workbook
from services import engineering_summary as ES
from ui import state
from ui.theme import kpi, chip, fmt, banner
from visualizations.dashboard import completion_chart, pqs_by_scenario

p, res, user = state.page_setup("Project Overview", "Create, load or import a project, then work through the engineering stages in order.")
summ = ES.build(p, res)

c = st.columns(5)
e = summ["electrical"]
c[0].markdown(kpi("Connected load", fmt(e["connected_kW"]), "kW", f"{fmt(e['connected_kVA'])} kVA arithmetic"), unsafe_allow_html=True)
c[1].markdown(kpi("Maximum demand", fmt(e["apparent_kVA_peak"]), "kVA", e["max_demand_scenario"] or "–"), unsafe_allow_html=True)
c[2].markdown(kpi("Transformers / feeders", f"{e['n_transformers']} / {e['n_feeders']}", "", f"max loading {fmt(e['max_tx_loading_pct'])} %"),
              unsafe_allow_html=True)
c[3].markdown(kpi("Daily auxiliary energy", fmt(e["daily_kWh"], 0), "kWh", "normal operating day"), unsafe_allow_html=True)
c[4].markdown(f'<div class="disp"><div class="t">Engineering disposition</div><div class="v">{summ["disposition"]}</div></div>',
              unsafe_allow_html=True)
st.write("")

t1, t2, t3 = st.tabs(["Project information", "Create / load / import", "Progress"])
with t1:
    with st.form("overview"):
        st.markdown("**Project identification** — fields marked 🔗 are the Design Basis inputs themselves (single source, no duplicate inputs).")
        db = p["design_basis"]
        vals = {}
        cols = st.columns(3)
        for i, (k, label) in enumerate(OVERVIEW_DB_LINKS):
            with cols[i % 3]:
                if k in ("BESS_MW", "BESS_MWh", "VLL"):
                    v = db.get(k)
                    vals[k] = st.number_input(f"🔗 {label}", value=float(v) if isinstance(v, (int, float)) else None, min_value=0.0,
                                              format="%.3f" if k != "VLL" else "%.0f", key=f"ov_{k}_{state.editor_ver()}")
                else:
                    vals[k] = st.text_input(f"🔗 {label}", value=str(db.get(k) or ""), key=f"ov_{k}_{state.editor_ver()}")
        st.markdown("**Additional project information** (informational only — no effect on calculations, approval A-3)")
        info = p["project_info"]
        ivals = {}
        cols = st.columns(3)
        for i, (k, label) in enumerate(PROJECT_INFO_FIELDS):
            with cols[i % 3]:
                ivals[k] = st.text_input(label, value=str(info.get(k) or ""), key=f"pi_{k}_{state.editor_ver()}")
        rev = st.text_input("Project file revision (used in the file name)", value=p["metadata"].get("revision", "Rev00"))
        dur = res["workbook"]["design_basis_derived"]["Duration_h"]
        st.caption(f"BESS duration (derived = MWh ÷ MW): {fmt(dur, 2)} h")
        if st.form_submit_button("Apply", type="primary"):
            for k, v in vals.items():
                db[k] = (int(v) if isinstance(v, float) and v.is_integer() else v) if not isinstance(v, str) else (v.strip() or None)
            for k, v in ivals.items():
                info[k] = v.strip() or None
            p["metadata"]["revision"] = rev.strip() or "Rev00"
            state.touch(); state.set_project(p); st.rerun()

with t2:
    a, b, c_ = st.columns(3)
    with a, st.container(border=True):
        st.markdown("**New project**")
        nm = st.text_input("Project name", key="new_name")
        ok = st.checkbox("I have saved the current project or want to discard it", key="new_confirm")
        if st.button("Create new project", disabled=not ok):
            state.set_project(new_project(nm, user)); st.session_state["unsaved"] = True; st.rerun()
    with b, st.container(border=True):
        st.markdown("**Upload existing project (JSON)**")
        f = st.file_uploader("Project file", type=["json"], key=f"up_json_{state.editor_ver()}")
        if f is not None:
            q, errs, warns = PS.from_json(f.getvalue())
            for e_ in errs:
                st.error(e_)
            for w in warns:
                st.warning(w)
            if q and st.button("Load this project (replaces the current session project)", type="primary"):
                state.set_project(q); st.session_state["unsaved"] = False; st.rerun()
    with c_, st.container(border=True):
        st.markdown("**Import v1.1 Excel workbook** (input cells only — A-7)")
        f2 = st.file_uploader("Workbook (.xlsx)", type=["xlsx"], key=f"up_x_{state.editor_ver()}")
        if f2 is not None:
            try:
                imp, rep = import_workbook(f2)
            except Exception as ex:   # corrupt / non-workbook file
                imp, rep = None, dict(errors=[f"Cannot read workbook: {ex}"], warnings=[], info=[])
            for e_ in rep["errors"]:
                st.error(e_)
            for w in rep["warnings"]:
                st.warning(w)
            for i_ in rep["info"]:
                st.caption(i_)
            if imp and st.button("Create project from these inputs", type="primary"):
                q = new_project(imp["design_basis"].get("Project") or "", user)
                for k in ("design_basis", "loads", "transformers", "feeders", "library", "categories", "profiles", "scenarios", "rating_series"):
                    if imp.get(k):
                        q[k] = imp[k]
                q["settings"].update({k: v for k, v in imp.get("settings", {}).items() if v is not None})
                q["metadata"]["import_report"] = rep
                state.set_project(q); st.session_state["unsaved"] = True; st.rerun()
        st.caption("Calculated cells are ignored; everything is recalculated by the engine. Missing inputs are reported, not defaulted.")

with t3:
    l, r = st.columns([1, 1.3])
    l.plotly_chart(completion_chart(summ["completion"]), use_container_width=True)
    with r:
        for it in summ["completion"]:
            st.markdown(f"{chip(it['status'] if it['status'] != 'Completed' else 'PASS', it['status'])} **{it['module']}** — "
                        f"<span class='small-muted'>{it['detail']}</span>", unsafe_allow_html=True)
    st.plotly_chart(pqs_by_scenario(res["workbook"]["scenarios"]["rows"]), use_container_width=True)
    st.markdown(banner("Workflow: Design Basis → Equipment Library → Load Schedule → Scenario Analysis → Transformer Sizing → "
                       "LV Distribution → Engineering Validation → Summary / Exports.", "info"), unsafe_allow_html=True)
