import streamlit as st

from calculations import reference_data as R
from data_models.project_schema import TABLE_INPUTS
from ui import state, tables
from ui.theme import banner

p, res, user = state.page_setup("Equipment Library", "Reusable generic equipment classes (sheet 2). Ratings are blank until entered from a "
                                "datasheet — none are invented. Manufacturer and DeratingNote are informational (A-4).")
L = R.LISTS
tab1, tab2, tab3 = st.tabs(["Equipment library", "Derived ratings & status", "Load categories"])
with tab1:
    opts = dict(Category=state.ids(p["categories"], "Category"), Phase=L["lstPhase"], RatingType=L["lstRType"], Nonlinear=L["lstNL"],
                Profile=state.ids(p["profiles"], "Profile"))
    tables.edit_table("library", p["library"], TABLE_INPUTS["library"], opts, "lib", height=520, id_col="LibID")
    st.markdown(banner("Per-equipment derating notes are references only. Transformer derating inputs and logic stay in the Design Basis "
                       "and Transformer Sizing (original workbook locations).", "info"), unsafe_allow_html=True)
with tab2:
    rows = []
    for l, s in zip(p["library"], res["workbook"]["library_status"]):
        rt, v, pf, eff = l.get("RatingType"), l.get("RatedValue"), l.get("PF"), l.get("Eff")
        kw = kva = None
        num = lambda x: isinstance(x, (int, float))
        if num(v):
            if rt == "kW_in":
                kw, kva = v, (v / pf if num(pf) and pf > 0 else None)
            elif rt == "kVA":
                kva, kw = v, (v * pf if num(pf) else None)
            elif rt == "kW_out" and num(eff) and eff > 0:
                kw = v / eff
                kva = kw / pf if num(pf) and pf > 0 else None
        rows.append(dict(LibID=l.get("LibID"), Category=l.get("Category"), RatingType=rt, RatedValue=v, kW_in_per_unit=kw,
                         kVA_per_unit=kva, Basis=("depends on applied voltage (load row)" if rt == "A" and num(v) else ""),
                         Manufacturer=l.get("Manufacturer"), Status=s))
    st.dataframe(tables.result_df(rows, list(rows[0].keys()) if rows else []), use_container_width=True, hide_index=True, height=520)
    st.caption("kW / kVA per unit are calculated (electrical input). A-rated items need the load-row voltage; shaft (kW_out) ratings use efficiency.")
with tab3:
    tables.edit_table("categories", p["categories"], TABLE_INPUTS["categories"], {}, "cat", height=420, id_col="Category")
