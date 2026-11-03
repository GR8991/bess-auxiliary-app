"""Email allowlist gate.

IMPORTANT - this is a BASIC EMAIL ALLOWLIST GATE, NOT IDENTITY AUTHENTICATION.
The application does not verify that the person typing an address owns it: anyone who knows an
approved address can enter.  Use it only for controlled preliminary engineering use.

Approved addresses are read from Streamlit Secrets:  st.secrets["access"]["allowed_emails"]
(never hardcoded).  Denial messages never reveal the list or whether an address exists.
"""
from __future__ import annotations

import time

GATE_DISCLOSURE = ("Access check: basic email allowlist gate. Your email address is compared with an approved list; "
                   "it does not verify your identity.")
MAX_ATTEMPTS = 5
LOCKOUT_SECONDS = 60
DEFAULT_TIMEOUT_MIN = 60


def normalize_email(s) -> str:
    return str(s or "").strip().lower()


def is_allowed(email: str, allowed) -> bool:
    e = normalize_email(email)
    if not e or "@" not in e:
        return False
    allowed_norm = {normalize_email(a) for a in (allowed or []) if normalize_email(a)}
    return e in allowed_norm


def load_allowlist(secrets) -> tuple[list, str | None]:
    """Return (allowed_emails, configuration_error)."""
    try:
        acc = secrets["access"]
        lst = list(acc["allowed_emails"])
    except Exception:
        return [], "Access list not configured. The administrator must set [access] allowed_emails in Streamlit Secrets."
    if not lst:
        return [], "Access list is empty. The administrator must add approved addresses in Streamlit Secrets."
    return lst, None


def session_timeout_minutes(secrets) -> int:
    try:
        return int(secrets["access"].get("session_timeout_minutes", DEFAULT_TIMEOUT_MIN))
    except Exception:
        return DEFAULT_TIMEOUT_MIN


class Gate:
    """Session logic, independent of Streamlit so it can be unit-tested. `state` is a dict-like
    (st.session_state in the app)."""

    def __init__(self, state, allowed, timeout_min=DEFAULT_TIMEOUT_MIN, clock=time.time):
        self.s, self.allowed, self.timeout, self.clock = state, allowed, timeout_min, clock

    def authorized(self) -> bool:
        if not self.s.get("auth_ok"):
            return False
        last = self.s.get("auth_last", 0)
        if self.timeout and self.clock() - last > self.timeout * 60:
            self.expire()
            return False
        self.s["auth_last"] = self.clock()
        return True

    def attempt(self, email) -> tuple[bool, str]:
        now = self.clock()
        until = self.s.get("auth_locked_until", 0)
        if now < until:
            return False, f"Too many attempts. Try again in {int(until - now) + 1} s."
        if is_allowed(email, self.allowed):
            self.s.update(auth_ok=True, auth_email=normalize_email(email), auth_last=now, auth_fail=0)
            return True, ""
        n = self.s.get("auth_fail", 0) + 1
        self.s["auth_fail"] = n
        if n >= MAX_ATTEMPTS:
            self.s["auth_locked_until"] = now + LOCKOUT_SECONDS
            self.s["auth_fail"] = 0
        return False, "Access denied. This address is not authorised for this application. Contact your administrator."

    def expire(self):
        keep = {"auth_notice": "Your session expired. Please sign in again. Unsaved project changes were cleared - "
                               "re-upload your latest project file."}
        for k in list(self.s.keys()):
            del self.s[k]
        self.s.update(keep)

    def logout(self):
        for k in list(self.s.keys()):
            del self.s[k]
        self.s["auth_notice"] = "You have signed out. Session data was cleared."


# ---------------------------------------------------------------- Streamlit glue
def get_gate():
    import streamlit as st
    allowed, err = load_allowlist(st.secrets) if _has_secrets(st) else ([], "Streamlit Secrets not found. Copy "
                                                                         ".streamlit/secrets.toml.example to .streamlit/secrets.toml.")
    return Gate(st.session_state, allowed, session_timeout_minutes(st.secrets) if _has_secrets(st) else DEFAULT_TIMEOUT_MIN), err


def _has_secrets(st):
    try:
        return "access" in st.secrets
    except Exception:
        return False


def require_auth():
    """Call at the top of every page: stops rendering for an unauthorised session."""
    import streamlit as st
    gate, _ = get_gate()
    if not gate.authorized():
        st.error("Your session is not authorised. Please sign in from the application home page.")
        st.stop()
    return st.session_state.get("auth_email")
