"""BESS Auxiliary Engineering - Streamlit entry point.

Run:  streamlit run app.py
Access: basic email allowlist gate (st.secrets["access"]["allowed_emails"]) - NOT identity verification.
Storage: session only + JSON download/upload. No server-side persistence.
"""
from __future__ import annotations

import streamlit as st

st.set_page_config(page_title="BESS Auxiliary Engineering", page_icon="⚡", layout="wide", initial_sidebar_state="expanded")

from authentication.access_control import get_gate, GATE_DISCLOSURE  # noqa: E402
from calculations.constants import ENGINE_VERSION, SOURCE_WORKBOOK_VERSION  # noqa: E402
from data_models.project_schema import APP_VERSION  # noqa: E402
from services import project_storage as PS  # noqa: E402
from ui import state  # noqa: E402
from ui.theme import CSS, LOGIN_CSS, BESS_SVG, logo_html, chip  # noqa: E402


def branding():
    try:
        b = st.secrets.get("branding", {})
    except Exception:
        b = {}
    return b.get("company_name", "SunStripe"), b.get("logo_path", "assets/logo.svg")


def login_page(gate, cfg_error):
    company, logo = branding()
    st.markdown(LOGIN_CSS, unsafe_allow_html=True)
    left, right = st.columns([1.25, 1], gap="large")
    with left:
        st.markdown(f'<div class="brand">{logo_html(company, logo, 30)}<span>{company} · Power Systems Engineering</span></div>'
                    '<div class="login-hero"><h1>BESS Auxiliary Load &amp;<br/>Transformer Sizing</h1>'
                    '<p>Engineering workspace for auxiliary load schedules, operating scenarios, auxiliary transformer sizing, '
                    'LV distribution screening and validation - calculation engine reproducing workbook v1.1.</p></div>',
                    unsafe_allow_html=True)
        st.markdown(BESS_SVG, unsafe_allow_html=True)
        st.markdown('<p class="feat"><b>Deterministic engine</b> · 119,876 workbook-equivalence comparisons executed</p>'
                    '<p class="feat"><b>Traceable</b> · every value tagged as input, default, derived or calculated</p>'
                    '<p class="feat"><b>Your data stays with you</b> · projects live in your session and your JSON files</p>',
                    unsafe_allow_html=True)
    with right, st.container(border=True):
        st.markdown('<h3 style="margin-top:0">Sign in</h3><p class="small-muted">Enter your approved company email address.</p>',
                    unsafe_allow_html=True)
        notice = st.session_state.pop("auth_notice", None)
        if notice:
            st.info(notice)
        if cfg_error:
            st.error(cfg_error)
        with st.form("login", clear_on_submit=False, border=False):
            email = st.text_input("Email address", placeholder="name@company.com", autocomplete="email")
            ok = st.form_submit_button("Continue", type="primary", use_container_width=True, disabled=bool(cfg_error))
        if ok:
            good, msg = gate.attempt(email)
            if good:
                st.rerun()
            else:
                st.error(msg)
        st.markdown(f'<div class="disclose">🔒 {GATE_DISCLOSURE} Anyone who knows an approved address can enter. '
                    'No passwords are requested or stored.</div>', unsafe_allow_html=True)
        st.caption(f"Application {APP_VERSION} · engine {ENGINE_VERSION} · workbook v{SOURCE_WORKBOOK_VERSION}")


def sidebar(user):
    company, logo = branding()
    p = state.project()
    res = state.results()
    with st.sidebar:
        st.markdown(f'<div style="display:flex;gap:10px;align-items:center;margin-bottom:6px">{logo_html(company, logo, 26)}'
                    f'<div><b>BESS Auxiliary Engineering</b><br/><span class="small-muted">{company}</span></div></div>',
                    unsafe_allow_html=True)
        name = p["design_basis"].get("Project") or "Untitled project"
        st.markdown(f"**{name}**  \n<span class='small-muted'>{p['metadata'].get('revision', '')} · updated "
                    f"{p['metadata'].get('updated', '')[:16].replace('T', ' ')} UTC</span>", unsafe_allow_html=True)
        st.markdown(chip(res["application"]["disposition"]), unsafe_allow_html=True)
        st.download_button("💾 Save project (JSON)", data=PS.to_json(p, res), file_name=PS.filename(p), mime="application/json",
                           use_container_width=True, type="primary", on_click=lambda: st.session_state.update(unsaved=False))
        if st.session_state.get("unsaved"):
            st.caption("⚠️ Unsaved changes — download the project file before leaving.")
        else:
            st.caption("Projects are not stored on the server. Keep your latest JSON file.")
        st.divider()
        st.markdown(f"<span class='small-muted'>Signed in: {user}</span>", unsafe_allow_html=True)
        if st.button("Log out", use_container_width=True):
            gate, _ = get_gate()
            gate.logout()
            st.rerun()


def main():
    gate, cfg_error = get_gate()
    if not gate.authorized():
        login_page(gate, cfg_error)
        st.stop()
    st.markdown(CSS, unsafe_allow_html=True)
    user = st.session_state.get("auth_email", "")
    state.project()
    pages = {
        "Project": [st.Page("pages/01_Project_Overview.py", title="Project Overview", icon=":material/dashboard:", default=True)],
        "Engineering inputs": [
            st.Page("pages/02_Design_Basis.py", title="Design Basis", icon=":material/tune:"),
            st.Page("pages/03_Equipment_Library.py", title="Equipment Library", icon=":material/inventory_2:"),
            st.Page("pages/04_Load_Schedule.py", title="Auxiliary Load Schedule", icon=":material/table_chart:")],
        "Analysis": [
            st.Page("pages/05_Scenario_Analysis.py", title="Scenario Analysis", icon=":material/stacked_bar_chart:"),
            st.Page("pages/06_Transformer_Sizing.py", title="Transformer Sizing", icon=":material/electric_bolt:"),
            st.Page("pages/07_LV_Distribution.py", title="LV Distribution", icon=":material/account_tree:"),
            st.Page("pages/08_Engineering_Validation.py", title="Engineering Validation", icon=":material/fact_check:")],
        "Outputs": [
            st.Page("pages/09_Engineering_Summary.py", title="Engineering Summary", icon=":material/summarize:"),
            st.Page("pages/10_AI_Agent_Handoff.py", title="AI Agent Handoff", icon=":material/hub:"),
            st.Page("pages/11_Reports_Exports.py", title="Reports and Exports", icon=":material/download:"),
            st.Page("pages/12_Settings.py", title="Application Settings", icon=":material/settings:")],
    }
    nav = st.navigation(pages)
    sidebar(user)
    nav.run()


main()
