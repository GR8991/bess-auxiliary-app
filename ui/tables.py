"""Editable engineering tables (st.data_editor) with explicit Apply, deletion confirmation,
row duplication, and lossless round-trip of fields that are not displayed (`_extra` etc.)."""
from __future__ import annotations

import math

import pandas as pd
import streamlit as st

from data_models.project_schema import NUMERIC
from ui.column_help import HELP
from ui import state

EXPRESSION_UI = {"loads": {"QtyInst", "QtyDuty"}}   # A-6: text cells accept '=DB_N_Encl/2' etc.


def _py(v):
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if hasattr(v, "item"):
        v = v.item()
    if isinstance(v, float) and v.is_integer() and abs(v) < 1e15:
        return int(v) if v == int(v) else v
    if isinstance(v, str):
        v = v.strip()
        return v if v != "" else None
    return v


def _parse_expr_cell(v):
    v = _py(v)
    if isinstance(v, str) and not v.startswith("="):
        try:
            f = float(v)
            return int(f) if f.is_integer() else f
        except ValueError:
            return v
    return v


def to_df(rows, cols, table):
    data = []
    for r in rows:
        d = {}
        for c in cols:
            v = r.get(c)
            if c in EXPRESSION_UI.get(table, ()):
                v = None if v is None else str(v)
            d[c] = v
        data.append(d)
    df = pd.DataFrame(data, columns=cols)
    for c in cols:
        if c in NUMERIC.get(table, set()) and c not in EXPRESSION_UI.get(table, ()):
            df[c] = pd.to_numeric(df[c], errors="coerce")
        else:
            df[c] = df[c].astype("object")
    return df


def from_df(df, cols, table, original):
    out = []
    for idx, row in df.iterrows():
        base = {}
        if isinstance(idx, (int,)) and 0 <= idx < len(original):
            base = {k: v for k, v in original[idx].items() if k not in cols}     # keep non-displayed fields (_extra ...)
        d = dict(base)
        for c in cols:
            v = row.get(c)
            d[c] = _parse_expr_cell(v) if c in EXPRESSION_UI.get(table, ()) else _py(v)
        if any(d.get(c) is not None for c in cols):
            out.append(d)
    return out


def column_config(table, cols, options: dict, widths: dict | None = None):
    cfg = {}
    help_ = HELP.get(table, {})
    for c in cols:
        h = help_.get(c, "")
        if c in options:
            cfg[c] = st.column_config.SelectboxColumn(c, options=[o for o in options[c] if o not in (None, "")], help=h)
        elif c in NUMERIC.get(table, set()) and c not in EXPRESSION_UI.get(table, ()):
            cfg[c] = st.column_config.NumberColumn(c, help=h, format="%.4g")
        else:
            cfg[c] = st.column_config.TextColumn(c, help=h + (" — accepts a number or an Excel-style formula such as =DB_N_Encl/2"
                                                              if c in EXPRESSION_UI.get(table, ()) else ""))
    return cfg


def edit_table(table, rows, cols, options, key, height=420, id_col=None):
    """Render an editable table. Returns True when changes were applied to the project."""
    p = state.project()
    df = to_df(rows, cols, table)
    edited = st.data_editor(df, key=f"{key}_{state.editor_ver()}", num_rows="dynamic", use_container_width=True, height=height,
                            column_config=column_config(table, cols, options), hide_index=False)
    new_rows = from_df(edited, cols, table, rows)
    changed = new_rows != rows
    applied = False
    c1, c2, c3 = st.columns([1.4, 1.4, 3])
    removed = len(rows) - sum(1 for i in edited.index if isinstance(i, int) and i < len(rows))
    if changed:
        with c3:
            st.markdown('<div class="banner warn">Unapplied table edits — press <b>Apply changes</b> to recalculate.</div>',
                        unsafe_allow_html=True)
        confirm = True
        if removed > 0:
            confirm = c2.checkbox(f"Confirm deletion of {removed} row(s)", key=f"{key}_del_{state.editor_ver()}")
        if c1.button("Apply changes", key=f"{key}_apply", type="primary", disabled=not confirm):
            p[table] = new_rows
            state.touch()
            state.set_project(p)
            applied = True
            st.rerun()
    if rows and id_col:
        with st.expander("Duplicate a row"):
            labels = [f"{i + 1}: {r.get(id_col) or '(no ID)'}" for i, r in enumerate(rows)]
            sel = st.selectbox("Row to duplicate", labels, key=f"{key}_dupsel")
            if st.button("Duplicate", key=f"{key}_dup"):
                i = labels.index(sel)
                cp = dict(rows[i])
                if cp.get(id_col):
                    cp[id_col] = f"{cp[id_col]}-COPY"
                p[table] = rows[: i + 1] + [cp] + rows[i + 1:]
                state.touch()
                state.set_project(p)
                st.rerun()
    return applied


def result_df(rows, cols, fmt_cols=None):
    from calculations.excel_semantics import XLError
    data = []
    for r in rows:
        d = {}
        for c in cols:
            v = r.get(c)
            if isinstance(v, XLError):
                v = v.code
            if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
                v = None
            d[c] = v
        data.append(d)
    df = pd.DataFrame(data, columns=cols)
    for c in cols:
        if df[c].map(lambda x: isinstance(x, str)).any():
            df[c] = df[c].astype("object").where(df[c].notna(), "")
    return df
