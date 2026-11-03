import json

import streamlit as st

from data_models.handoff_schema import HANDOFF_SCHEMA
from services import json_export as JE
from ui import state
from ui.theme import banner, kpi

p, res, user = state.page_setup("AI Agent Handoff", "Structured engineering data package for the separate Power System BOP Engineer AI Agent.")
st.markdown(banner("The calculation engine does not depend on any AI service. This package only transfers deterministic results; every value "
                   "is tagged <b>user_input / workbook_default / derived / calculated / missing</b>. Missing inputs are listed, never filled.",
                   "info"), unsafe_allow_html=True)
pkg = JE.build_handoff(p, res, user)
try:
    JE.validate_handoff(pkg)
    valid, msg = True, "Package validates against the handoff JSON Schema."
except Exception as ex:
    valid, msg = False, f"Schema validation failed: {ex}"
c = st.columns(4)
c[0].markdown(kpi("Schema", pkg["handoff_schema_version"], "", "valid" if valid else "INVALID"), unsafe_allow_html=True)
c[1].markdown(kpi("Disposition", pkg["overall_disposition"]["status"][:28]), unsafe_allow_html=True)
c[2].markdown(kpi("Outstanding data items", len(pkg["outstanding_data_requirements"])), unsafe_allow_html=True)
c[3].markdown(kpi("Findings", len(pkg["validation"]["findings"])), unsafe_allow_html=True)
(st.success if valid else st.error)(msg)
a, b = st.columns(2)
name = f"{(p['design_basis'].get('Project') or 'Project')}_BOP_handoff_{p['metadata'].get('revision', 'Rev00')}.json".replace(" ", "_")
a.download_button("Download handoff package (JSON)", JE.to_bytes(pkg), file_name=name, mime="application/json", type="primary",
                  use_container_width=True, disabled=not valid)
b.download_button("Download handoff JSON Schema", json.dumps(HANDOFF_SCHEMA, indent=2).encode(), file_name="bess_aux_handoff_schema_1.0.json",
                  mime="application/json", use_container_width=True)
with st.expander("Package sections"):
    st.json({k: (f"{len(v)} items" if isinstance(v, list) else (list(v)[:12] if isinstance(v, dict) else v)) for k, v in pkg.items()},
            expanded=True)
with st.expander("Preview: identification, disposition and outstanding data"):
    st.json({"identification": pkg["identification"], "overall_disposition": pkg["overall_disposition"],
             "outstanding_data_requirements": pkg["outstanding_data_requirements"]}, expanded=False)
