"""Plotly figures built ONLY from engine outputs. Common layout = engineering dark theme."""
from __future__ import annotations

import plotly.graph_objects as go

from ui.theme import CHART, TEXT, MUTED, LINE, PANEL, STATUS_COLORS, BLUE

SCEN_LABEL = {"NORM": "Normal", "MAXCOOL": "Max cooling", "MAXSIM": "Max simultaneous", "STBY": "Standby", "MAINT": "Maintenance",
              "EMERG": "Emergency", "BLACK": "Black-start"}


def layout(fig, title="", height=360, legend=True):
    fig.update_layout(title=dict(text=title, x=0.01, font=dict(size=14, color=TEXT)), height=height, paper_bgcolor="rgba(0,0,0,0)",
                      plot_bgcolor="rgba(0,0,0,0)", font=dict(color=TEXT, size=12), margin=dict(l=10, r=10, t=46, b=10),
                      showlegend=legend, legend=dict(orientation="h", y=-0.18, font=dict(color=MUTED)), colorway=CHART)
    fig.update_xaxes(gridcolor=LINE, zerolinecolor=LINE, color=MUTED)
    fig.update_yaxes(gridcolor=LINE, zerolinecolor=LINE, color=MUTED)
    return fig


def empty(title, msg="No calculated data yet"):
    fig = go.Figure()
    fig.add_annotation(text=msg, x=0.5, y=0.5, xref="paper", yref="paper", showarrow=False, font=dict(color=MUTED))
    fig.update_xaxes(visible=False)
    fig.update_yaxes(visible=False)
    return layout(fig, title, 260, False)


def category_donut(cats, value_key="Conn_kW", title="Connected load by category (kW)"):
    items = [(c["Category"], c.get(value_key) or 0) for c in cats if (c.get(value_key) or 0) > 0]
    if not items:
        return empty(title)
    fig = go.Figure(go.Pie(labels=[i[0] for i in items], values=[i[1] for i in items], hole=0.58, sort=True,
                           textinfo="percent", hovertemplate="%{label}<br>%{value:,.1f} kW<br>%{percent}<extra></extra>"))
    return layout(fig, title, 360)


def pqs_by_scenario(rows, title="Active, reactive and apparent power by scenario"):
    x = [SCEN_LABEL.get(r["Code"], r["Code"]) for r in rows]
    fig = go.Figure()
    fig.add_bar(name="P (kW)", x=x, y=[r["P_counted"] for r in rows])
    fig.add_bar(name="Q (kvar, signed)", x=x, y=[r["Q"] for r in rows])
    fig.add_bar(name="S (kVA, vector)", x=x, y=[r["S"] for r in rows])
    fig.add_scatter(name="Transient S incl. largest start (kVA)", x=x, y=[r["S_transient"] for r in rows], mode="markers",
                    marker=dict(symbol="diamond", size=10, color=STATUS_COLORS["REVIEW"]))
    fig.update_layout(barmode="group")
    return layout(fig, title, 380)


def tx_loading(tx, title="Transformer loading vs evaluated rating"):
    if not tx:
        return empty(title, "No transformers defined")
    ids_ = [t["TX_ID"] for t in tx]
    fig = go.Figure()
    fig.add_bar(name="Evaluated rating (kVA)", y=ids_, x=[t["Eval_kVA"] or 0 for t in tx], orientation="h",
                marker=dict(color="rgba(127,178,255,0.18)", line=dict(color=BLUE, width=1)))
    fig.add_bar(name="Derated capacity (kVA)", y=ids_, x=[t["Cap_derated"] or 0 for t in tx], orientation="h",
                marker=dict(color="rgba(95,208,198,0.22)", line=dict(color="#5FD0C6", width=1)))
    fig.add_bar(name="Design demand (kVA)", y=ids_, x=[t["Design_kVA"] or 0 for t in tx], orientation="h",
                marker=dict(color=[STATUS_COLORS["FAIL"] if (t["LoadDer_pct"] or 0) > 100 else BLUE for t in tx]),
                text=[f"{t['Load_pct']:.0f} %" for t in tx], textposition="outside")
    fig.update_layout(barmode="overlay")
    fig.update_yaxes(autorange="reversed")
    return layout(fig, title, max(260, 60 + 42 * len(tx)))


def tx_utilisation(tx, scen_codes, title="Transformer utilisation by scenario (% of evaluated rating)"):
    if not tx:
        return empty(title, "No transformers defined")
    fig = go.Figure(go.Heatmap(z=[[t["U_" + s] for s in scen_codes] for t in tx], x=[SCEN_LABEL.get(s, s) for s in scen_codes],
                               y=[t["TX_ID"] for t in tx], colorscale=[[0, "#12233A"], [0.6, BLUE], [0.8, "#E0A030"], [1, "#E5484D"]],
                               zmin=0, zmax=120, texttemplate="%{z:.0f}", colorbar=dict(title="%")))
    fig.update_yaxes(autorange="reversed")
    return layout(fig, title, max(260, 80 + 34 * len(tx)), legend=False)


def feeder_loading(fdr, title="Feeder loading (Ib ÷ In) — screening"):
    items = [(f["FeederID"], f["Load_pct"]) for f in fdr if isinstance(f.get("Load_pct"), (int, float))]
    if not items:
        return empty(title, "No feeder with a protective-device rating entered")
    fig = go.Figure(go.Bar(x=[i[0] for i in items], y=[i[1] for i in items],
                           marker=dict(color=[STATUS_COLORS["FAIL"] if i[1] > 100 else BLUE for i in items]),
                           text=[f"{i[1]:.0f} %" for i in items], textposition="outside"))
    fig.add_hline(y=100, line=dict(color=STATUS_COLORS["REVIEW"], dash="dash"), annotation_text="100 % of In")
    return layout(fig, title, 340, legend=False)


def phase_balance(rows, id_key, title):
    if not rows:
        return empty(title)
    x = [r[id_key] for r in rows]
    fig = go.Figure()
    for ph in ("L1", "L2", "L3"):
        fig.add_bar(name=ph, x=x, y=[r[f"{ph}_kVA"] for r in rows])
    fig.update_layout(barmode="group")
    return layout(fig, title, 340)


def status_bar(counts, title="Validation status"):
    order = ["PASS", "FAIL", "REVIEW", "INCOMPLETE", "NOT VERIFIED", "NOT APPLICABLE"]
    fig = go.Figure(go.Bar(x=[counts.get(k, 0) for k in order], y=order, orientation="h",
                           marker=dict(color=[STATUS_COLORS[k] for k in order]), text=[counts.get(k, 0) for k in order],
                           textposition="outside"))
    fig.update_yaxes(autorange="reversed")
    return layout(fig, title, 300, legend=False)


def completion_chart(items, title="Input completion by module"):
    score = {"Completed": 3, "Partially completed": 2, "Requires engineering review": 1.5, "Missing mandatory inputs": 1}
    col = {"Completed": STATUS_COLORS["PASS"], "Partially completed": STATUS_COLORS["REVIEW"],
           "Requires engineering review": STATUS_COLORS["NOT VERIFIED"], "Missing mandatory inputs": STATUS_COLORS["FAIL"]}
    fig = go.Figure(go.Bar(y=[i["module"] for i in items], x=[score[i["status"]] for i in items], orientation="h",
                           marker=dict(color=[col[i["status"]] for i in items]), text=[i["status"] for i in items],
                           textposition="inside", insidetextanchor="start", hovertext=[i["detail"] for i in items]))
    fig.update_xaxes(visible=False, range=[0, 3])
    fig.update_yaxes(autorange="reversed")
    return layout(fig, title, 320, legend=False)
