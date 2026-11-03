"""Design system: palette, CSS, status chips, login visuals (all inline - no external assets required)."""
from __future__ import annotations

import base64
import html
import os

NAVY = "#0B1626"
NAVY_2 = "#12233A"
PANEL = "#16283F"
LINE = "#26405F"
BLUE = "#3D8BFD"
BLUE_2 = "#7FB2FF"
TEXT = "#E6EDF5"
MUTED = "#9FB0C4"
STATUS_COLORS = {"PASS": "#2FA36B", "FAIL": "#E5484D", "REVIEW": "#E0A030", "INCOMPLETE": "#E0A030", "NOT VERIFIED": "#9B87F5",
                 "NOT APPLICABLE": "#6B7A8F", "High": "#E5484D", "Medium": "#E0A030", "Low": "#7FB2FF"}
CHART = [BLUE, "#5FD0C6", "#9B87F5", "#E0A030", "#7FB2FF", "#2FA36B", "#E5484D", "#C9D3DF", "#3A6EA5", "#B8E07A"]
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

GRID_SVG = ("<svg xmlns='http://www.w3.org/2000/svg' width='64' height='64'><path d='M64 0H0V64' fill='none' stroke='%2326405F' "
            "stroke-width='0.6' opacity='0.55'/><circle cx='0' cy='0' r='1.6' fill='%233D8BFD' opacity='0.5'/></svg>")

CSS = f"""
<style>
:root {{ --navy:{NAVY}; --panel:{PANEL}; --line:{LINE}; --blue:{BLUE}; --text:{TEXT}; --muted:{MUTED}; }}
html, body, [class*="css"] {{ font-family: 'Inter', 'Segoe UI', system-ui, -apple-system, sans-serif; }}
.block-container {{ padding-top: 1.6rem; padding-bottom: 3rem; max-width: 1500px; }}
h1, h2, h3 {{ letter-spacing: -0.01em; }}
h1 {{ font-weight: 650; font-size: 1.75rem !important; }}
[data-testid="stSidebar"] {{ background: linear-gradient(180deg, {NAVY_2} 0%, {NAVY} 100%); border-right: 1px solid {LINE}; }}
.eng-card {{ background:{PANEL}; border:1px solid {LINE}; border-radius:10px; padding:14px 16px; margin-bottom:10px; }}
.eng-kpi {{ background:{PANEL}; border:1px solid {LINE}; border-radius:10px; padding:12px 14px; height:100%; }}
.eng-kpi .lab {{ color:{MUTED}; font-size:0.78rem; text-transform:uppercase; letter-spacing:.04em; }}
.eng-kpi .val {{ color:{TEXT}; font-size:1.45rem; font-weight:650; margin-top:2px; }}
.eng-kpi .unit {{ color:{MUTED}; font-size:0.85rem; font-weight:500; margin-left:4px; }}
.eng-kpi .sub {{ color:{MUTED}; font-size:0.75rem; margin-top:2px; }}
.chip {{ display:inline-block; padding:2px 9px; border-radius:999px; font-size:0.74rem; font-weight:600; letter-spacing:.02em;
        border:1px solid rgba(255,255,255,.08); }}
.banner {{ border-radius:10px; padding:10px 14px; margin:6px 0 12px 0; border:1px solid; font-size:0.9rem; }}
.banner.warn {{ background:rgba(224,160,48,.10); border-color:rgba(224,160,48,.45); }}
.banner.info {{ background:rgba(61,139,253,.10); border-color:rgba(61,139,253,.45); }}
.banner.bad {{ background:rgba(229,72,77,.10); border-color:rgba(229,72,77,.45); }}
.disp {{ border-radius:12px; padding:14px 18px; border:1px solid {LINE}; background:linear-gradient(90deg, rgba(61,139,253,.14), rgba(61,139,253,0)); }}
.disp .t {{ font-size:0.75rem; color:{MUTED}; text-transform:uppercase; letter-spacing:.06em; }}
.disp .v {{ font-size:1.25rem; font-weight:700; margin-top:2px; }}
.small-muted {{ color:{MUTED}; font-size:0.8rem; }}
.label-app {{ color:#9B87F5; font-weight:600; font-size:0.78rem; }}
div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] {{ border:1px solid {LINE}; border-radius:8px; }}
</style>
"""

LOGIN_CSS = f"""
<style>
[data-testid="stSidebar"], [data-testid="collapsedControl"], header[data-testid="stHeader"] {{ display:none; }}
.stApp {{ background-color:{NAVY};
  background-image: radial-gradient(1100px 600px at 85% 10%, rgba(61,139,253,.20), transparent 60%),
                    radial-gradient(900px 500px at 10% 95%, rgba(95,208,198,.10), transparent 60%),
                    url("data:image/svg+xml;utf8,{GRID_SVG}"); }}
.block-container {{ max-width: 1180px; padding-top: 6vh; }}
.login-hero h1 {{ font-size: 2.35rem !important; line-height:1.15; margin: 0 0 .4rem 0; color:{TEXT}; }}
.login-hero p {{ color:{MUTED}; font-size: 1.02rem; }}
.brand {{ display:flex; align-items:center; gap:10px; color:{TEXT}; font-weight:650; letter-spacing:.06em; text-transform:uppercase;
          font-size:.85rem; margin-bottom:28px; }}
.login-card {{ background: rgba(18,35,58,.86); border:1px solid {LINE}; border-radius:14px; padding:26px 26px 10px 26px;
               box-shadow: 0 20px 60px rgba(0,0,0,.35); }}
.login-card h3 {{ margin-top:0; }}
.disclose {{ color:{MUTED}; font-size:.78rem; border-top:1px solid {LINE}; padding-top:10px; margin-top:6px; }}
.feat {{ color:{MUTED}; font-size:.9rem; margin:4px 0; }}
.feat b {{ color:{TEXT}; font-weight:600; }}
</style>
"""

# Generic, original line illustration: BESS enclosures, auxiliary transformer, MV line (no third-party artwork)
BESS_SVG = f"""
<svg viewBox="0 0 560 260" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="BESS enclosures and auxiliary transformer">
 <defs><linearGradient id="g" x1="0" x2="0" y1="0" y2="1"><stop offset="0" stop-color="{BLUE}" stop-opacity=".35"/>
  <stop offset="1" stop-color="{BLUE}" stop-opacity=".05"/></linearGradient></defs>
 <line x1="10" y1="222" x2="550" y2="222" stroke="{LINE}" stroke-width="2"/>
 <g stroke="{BLUE_2}" stroke-width="1.6" fill="url(#g)">
  <rect x="24" y="128" width="150" height="94" rx="4"/><rect x="190" y="128" width="150" height="94" rx="4"/>
 </g>
 <g stroke="{BLUE_2}" stroke-width="1" opacity=".7">
  {''.join(f'<line x1="{x}" y1="140" x2="{x}" y2="210"/>' for x in range(40, 170, 16))}
  {''.join(f'<line x1="{x}" y1="140" x2="{x}" y2="210"/>' for x in range(206, 336, 16))}
 </g>
 <g fill="none" stroke="#5FD0C6" stroke-width="1.6">
  <rect x="380" y="150" width="70" height="72" rx="4"/><circle cx="403" cy="134" r="12"/><circle cx="427" cy="134" r="12"/>
  <line x1="415" y1="40" x2="415" y2="122"/><line x1="470" y1="40" x2="470" y2="222"/><line x1="530" y1="40" x2="530" y2="222"/>
  <path d="M415 40 L470 40 L530 40"/><path d="M380 186 L340 186"/>
 </g>
 <g fill="{TEXT}" font-family="Inter, Segoe UI, sans-serif" font-size="11" opacity=".75">
  <text x="24" y="118">BESS enclosures</text><text x="372" y="244">Aux transformer</text><text x="466" y="32">MV</text>
 </g>
 <g fill="#5FD0C6">{''.join(f'<circle cx="{x}" cy="40" r="2.5"/>' for x in (415, 470, 530))}</g>
</svg>"""


def logo_html(company="SunStripe", logo_path=None, size=30):
    if logo_path:
        p = logo_path if os.path.isabs(logo_path) else os.path.join(ROOT, logo_path)
        if os.path.exists(p):
            ext = os.path.splitext(p)[1].lower().strip(".")
            mime = "image/svg+xml" if ext == "svg" else f"image/{ext}"
            b64 = base64.b64encode(open(p, "rb").read()).decode()
            return f'<img src="data:{mime};base64,{b64}" height="{size}" alt="logo"/>'
    return (f'<svg width="{size}" height="{size}" viewBox="0 0 32 32"><rect width="32" height="32" rx="7" fill="{BLUE}"/>'
            f'<path d="M9 19h6l-2 6 10-12h-6l2-6z" fill="#fff"/></svg>')


def chip(status, label=None):
    s = str(status or "")
    key = next((k for k in STATUS_COLORS if s.upper().startswith(k.upper())), None)
    if s.startswith(("ERR", "FAIL")):
        key = "FAIL"
    elif s.startswith(("WARN", "PROVISIONAL")):
        key = "REVIEW"
    elif s.startswith(("OK", "SELECTED - adequate")):
        key = "PASS"
    c = STATUS_COLORS.get(key, MUTED)
    return f'<span class="chip" style="background:{c}22;color:{c};border-color:{c}55">{html.escape(label or s)}</span>'


def kpi(label, value, unit="", sub=""):
    return (f'<div class="eng-kpi"><div class="lab">{html.escape(label)}</div><div class="val">{html.escape(str(value))}'
            f'<span class="unit">{html.escape(unit)}</span></div><div class="sub">{html.escape(sub)}</div></div>')


def banner(text, kind="info"):
    return f'<div class="banner {kind}">{text}</div>'


def fmt(v, nd=1, dash="–"):
    try:
        from calculations.excel_semantics import XLError
        if isinstance(v, XLError):
            return v.code
    except Exception:
        pass
    if v is None or v == "":
        return dash
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (int, float)):
        return f"{v:,.{nd}f}"
    return str(v)
