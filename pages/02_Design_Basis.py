import streamlit as st

from calculations import reference_data as R
from data_models.project_schema import DB_TEXT_KEYS, DB_LIST, DB_MANDATORY, DB_PERCENT
from ui import state
from ui.theme import chip, banner

p, res, user = state.page_setup("Design Basis", "Global electrical, environmental and sizing inputs (workbook sheet 1). "
                                "★ = mandatory. Percent inputs are entered in %, stored as fractions exactly like the workbook.")
db = p["design_basis"]
chk = res["workbook"]["design_basis_checks"]
eff = res["workbook"]["design_basis"]
d = res["workbook"]["design_basis_derived"]
st.markdown(banner(f"Input checks: <b>{d['DB_ErrCount']}</b> error(s), <b>{d['DB_RevCount']}</b> review item(s). "
                   "Derating coefficients are deliberately blank - obtain them from the applicable standard / manufacturer; "
                   "a missing coefficient is never treated as 'no derating confirmed' (F-04).",
                   "bad" if d["DB_ErrCount"] else ("warn" if d["DB_RevCount"] else "info")), unsafe_allow_html=True)

sections = []
for it in R.DB_ITEMS:
    if not sections or sections[-1][0] != it["section"]:
        sections.append((it["section"], []))
    sections[-1][1].append(it)

v = state.editor_ver()
for title, items in sections:
    with st.expander(title or "General", expanded=title is not None and title.startswith(("3.", "5.", "6."))):
        with st.form(f"db_{title}"):
            new = {}
            for it in items:
                k = it["key"]
                label = f"{it['label']}{' ★' if k in DB_MANDATORY else ''}"
                unit = it["unit"] or ""
                note = it["note"] or ""
                cur = db.get(k)
                a, b = st.columns([3, 1.2])
                with a:
                    if it["derived"]:
                        st.text_input(f"{label} [{unit}] — derived", value="" if eff.get(k) in (None, "") else f"{eff[k]:.3f}",
                                      disabled=True, help=note, key=f"dbd_{k}_{v}")
                        continue
                    if isinstance(cur, str) and cur.startswith("="):
                        new[k] = st.text_input(f"{label} [{unit}] (formula)", value=cur, help=note, key=f"db_{k}_{v}")
                    elif k in DB_LIST:
                        opts = [""] + [str(o) for o in R.LISTS[DB_LIST[k]]]
                        sv = "" if cur is None else str(cur)
                        new[k] = st.selectbox(f"{label}", opts, index=opts.index(sv) if sv in opts else 0, help=note, key=f"db_{k}_{v}")
                    elif k in DB_TEXT_KEYS:
                        new[k] = st.text_input(label, value=str(cur or ""), help=note, key=f"db_{k}_{v}")
                    elif k in DB_PERCENT:
                        new[k] = st.number_input(f"{label} [%]", value=None if not isinstance(cur, (int, float)) else float(cur) * 100,
                                                 format="%.2f", help=note, key=f"db_{k}_{v}")
                    else:
                        new[k] = st.number_input(f"{label} [{unit}]", value=None if not isinstance(cur, (int, float)) else float(cur),
                                                 format="%.4g", help=note, key=f"db_{k}_{v}")
                with b:
                    c = chk.get("DB_" + k)
                    st.markdown("<div style='height:1.9rem'></div>" + (chip(c) if c else "<span class='small-muted'>no check</span>"),
                                unsafe_allow_html=True)
            if st.form_submit_button("Apply section", type="primary"):
                for k, val in new.items():
                    if k == "Freq" and val not in ("", None):
                        val = int(val)
                    elif isinstance(val, str):
                        val = val.strip() or None
                    elif isinstance(val, float):
                        val = val / 100 if k in DB_PERCENT else val
                        val = int(val) if float(val).is_integer() and k not in DB_PERCENT else val
                    db[k] = val
                state.touch(); state.set_project(p); st.rerun()
st.caption("Source of defaults: blank v1.1 template. Rating-basis defaults (40 °C / 1000 m) are commonly used normal service conditions "
           "and must be verified against the applicable standard edition and specification.")
