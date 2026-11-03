import streamlit as st

from authentication.access_control import GATE_DISCLOSURE
from calculations import reference_data as R
from calculations.constants import ENGINE_VERSION, SOURCE_WORKBOOK_VERSION, SC
from data_models.project_schema import APP_VERSION, SCHEMA_VERSION
from data_models.handoff_schema import HANDOFF_SCHEMA_VERSION
from ui import state
from ui.theme import banner

p, res, user = state.page_setup("Application Settings", "Project calculation settings and application information.")
s = p["settings"]
with st.form("settings"):
    st.markdown("**Project calculation settings** (workbook Summary / Transformer Sizing controls)")
    a, b, c = st.columns(3)
    days = a.number_input("Days per year for annual energy", min_value=0, max_value=366, value=int(s.get("days_per_year") or 365))
    brk = b.selectbox("Category breakdown scenario", SC, index=SC.index(s.get("breakdown_scenario", "MAXSIM")) if s.get("breakdown_scenario") in SC else 2)
    ct = c.selectbox("Centralized comparison: transformer type", R.LISTS["lstTXType"],
                     index=R.LISTS["lstTXType"].index(s.get("central_type")) if s.get("central_type") in R.LISTS["lstTXType"] else 0)
    st.markdown("**User rating series (kVA)** — used when Design Basis 'Standard rating series' = User list. Verify against manufacturer offer.")
    us = st.text_area("Comma-separated ratings", ", ".join(str(v) for v in p["rating_series"].get("User list", []) if isinstance(v, (int, float))))
    if st.form_submit_button("Apply settings", type="primary"):
        vals = []
        for t in us.replace(";", ",").split(","):
            t = t.strip()
            if t:
                try:
                    vals.append(float(t) if "." in t else int(t))
                except ValueError:
                    st.error(f"Not a number: {t}")
                    st.stop()
        s.update(days_per_year=int(days), breakdown_scenario=brk, central_type=ct)
        p["rating_series"]["User list"] = sorted(vals) + [None] * max(0, 40 - len(vals))
        state.touch(); state.set_project(p); st.rerun()
st.markdown("**Application information**")
st.table({"Item": ["Application version", "Calculation engine", "Source workbook", "Project schema", "Handoff schema", "Reference baseline"],
          "Value": [APP_VERSION, ENGINE_VERSION, f"v{SOURCE_WORKBOOK_VERSION}", SCHEMA_VERSION, HANDOFF_SCHEMA_VERSION,
                    "LibreOffice-recalculated v1.1 workbook (Excel desktop not yet verified)"]})
st.markdown(banner(f"<b>Access control:</b> {GATE_DISCLOSURE} Anyone who knows an approved address can enter.", "warn"), unsafe_allow_html=True)
st.markdown(banner("<b>Data protection:</b> projects exist only in your browser session and in the JSON files you download. Nothing is "
                   "written to a server database or disk, no project data is logged, and sessions are isolated per user. Download your "
                   "latest project file before closing the session.", "info"), unsafe_allow_html=True)
