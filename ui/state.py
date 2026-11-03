"""Session state: the active project lives ONLY in st.session_state (never on the server's disk,
never in a shared cache, never logged).  Results are recalculated by the engine whenever inputs change."""
from __future__ import annotations

import hashlib
import json

import streamlit as st

from calculations import engine
from data_models.project_schema import new_project, engine_input, now_iso
from authentication.access_control import require_auth
from ui.theme import CSS


def project() -> dict:
    if "project" not in st.session_state:
        st.session_state["project"] = new_project("", st.session_state.get("auth_email", ""))
        st.session_state["editor_ver"] = 0
    return st.session_state["project"]


def set_project(p: dict, reset_editors=True):
    st.session_state["project"] = p
    if reset_editors:
        st.session_state["editor_ver"] = st.session_state.get("editor_ver", 0) + 1
    st.session_state.pop("_res_hash", None)


def touch():
    """Record an input change (marks unsaved changes)."""
    p = project()
    p["metadata"]["updated"] = now_iso()
    st.session_state["unsaved"] = True
    st.session_state.pop("_res_hash", None)


def editor_ver() -> int:
    return st.session_state.get("editor_ver", 0)


def results() -> dict:
    p = project()
    h = hashlib.sha256(json.dumps(engine_input(p), sort_keys=True, default=str).encode()).hexdigest()
    if st.session_state.get("_res_hash") != h:
        st.session_state["_res"] = engine.run(engine_input(p))
        st.session_state["_res_hash"] = h
    return st.session_state["_res"]


def page_setup(title: str, subtitle: str = ""):
    """Common page header + authorisation guard. Returns (project, results, user)."""
    user = require_auth()
    st.markdown(CSS, unsafe_allow_html=True)
    st.title(title)
    if subtitle:
        st.caption(subtitle)
    return project(), results(), user


def ids(rows, key):
    return [str(r.get(key)) for r in rows or [] if r.get(key) not in (None, "")]
