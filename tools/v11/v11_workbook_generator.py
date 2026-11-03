"""BESS Auxiliary Load & Auxiliary Transformer Sizing Workbook - generator.
Builds a macro-free .xlsx with Excel Tables, structured references, named inputs,
data validation and traceable formulas.  Compatible with Excel 2010+ (no dynamic arrays).
"""
import re
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo, TableFormula
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter as CL

VERSION = "1.1"
N_VAL = 26
SC = ["NORM", "MAXCOOL", "MAXSIM", "STBY", "MAINT", "EMERG", "BLACK"]
SC_NAME = ["Normal BESS operation", "Maximum cooling demand", "Maximum simultaneous demand",
           "Standby (BESS idle, energised)", "Maintenance operation", "Emergency operation",
           "Black-start / restoration"]

# ---------------------------------------------------------------- styles
F_IN = PatternFill("solid", fgColor="FFF2CC")
F_CALC = PatternFill("solid", fgColor="EDEDED")
F_HDR = PatternFill("solid", fgColor="1F3864")
F_SUB = PatternFill("solid", fgColor="D9E1F2")
F_KPI = PatternFill("solid", fgColor="DDEBF7")
FN_IN = Font(color="0000CC")
FN_HDR = Font(color="FFFFFF", bold=True)
FN_B = Font(bold=True)
FN_T = Font(bold=True, size=14, color="1F3864")
FN_NOTE = Font(italic=True, size=9, color="595959")
THIN = Side(style="thin", color="BFBFBF")
BRD = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
RED = PatternFill("solid", fgColor="F8CBAD")
AMB = PatternFill("solid", fgColor="FFE699")
GRN = PatternFill("solid", fgColor="C6EFCE")
BLU = PatternFill("solid", fgColor="DDEBF7")


def this_row(tbl, f):
    return re.sub(r"\[@([A-Za-z0-9_]+)\]", lambda m: f"{tbl}[[#This Row],[{m.group(1)}]]", f)


def title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = FN_T
    if sub:
        ws["A2"] = sub
        ws["A2"].font = FN_NOTE


def add_name(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def status_cf(ws, rng):
    first = rng.split(":")[0]
    col = re.match(r"[A-Z]+", first).group(0)
    row = re.search(r"\d+", first).group(0)
    c = f"{col}{row}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({c},3)="ERR"'], fill=RED))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({c},4)="FAIL"'], fill=RED))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({c},4)="WARN"'], fill=AMB))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({c},6)="REVIEW"'], fill=AMB))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({c},11)="PROVISIONAL"'], fill=AMB))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({c},4)="PASS"'], fill=GRN))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({c},2)="OK"'], fill=GRN))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({c},8)="SELECTED"'], fill=GRN))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({c},4)="INFO"'], fill=BLU))


def make_table(wb, ws, name, cols, rows, top, left=1, nblank=0, style="TableStyleLight9"):
    """cols: list of dict(h, kind 'in'|'calc', f, fmt, w, unit, dv)
    rows: list of dict header->value (inputs).  Returns (first_data_row, last_data_row)."""
    hdr = top
    n = max(len(rows) + nblank, 1)
    for j, c in enumerate(cols):
        col = left + j
        cell = ws.cell(row=hdr - 1, column=col, value=c.get("unit", ""))
        cell.font = FN_NOTE
        cell.alignment = WRAP
        h = ws.cell(row=hdr, column=col, value=c["h"])
        h.fill = F_HDR if c["kind"] == "in" else PatternFill("solid", fgColor="44546A")
        h.font = FN_HDR
        h.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[CL(col)].width = c.get("w", 11)
    for i in range(n):
        r = hdr + 1 + i
        data = rows[i] if i < len(rows) else {}
        for j, c in enumerate(cols):
            cell = ws.cell(row=r, column=left + j)
            if c["kind"] == "in":
                v = data.get(c["h"], c.get("default"))
                if v is not None and v != "":
                    cell.value = v
                cell.fill = F_IN
                cell.font = FN_IN
            else:
                cell.value = "=" + this_row(name, c["f"])
                cell.fill = F_CALC
            if c.get("fmt"):
                cell.number_format = c["fmt"]
            cell.border = BRD
    last = hdr + n
    ref = f"{CL(left)}{hdr}:{CL(left + len(cols) - 1)}{last}"
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=False)
    t._initialise_columns()
    for j, c in enumerate(cols):
        t.tableColumns[j].name = c["h"]
        if c["kind"] == "calc":
            t.tableColumns[j].calculatedColumnFormula = TableFormula(attr_text=this_row(name, c["f"]))
    ws.add_table(t)
    # data validation over table rows plus spare rows for expansion
    for j, c in enumerate(cols):
        if c.get("dv"):
            dv = DataValidation(type="list", formula1=c["dv"], allow_blank=True,
                                showErrorMessage=True, errorTitle="Invalid entry",
                                error="Select a value from the drop-down list (see Lists sheet).")
            ws.add_data_validation(dv)
            dv.add(f"{CL(left + j)}{hdr + 1}:{CL(left + j)}{last + 300}")
        if c.get("dvnum"):
            lo, hi = c["dvnum"]
            dv = DataValidation(type="decimal", operator="between", formula1=str(lo), formula2=str(hi),
                                allow_blank=True, showErrorMessage=True, errorTitle="Out of range",
                                error=f"Enter a number between {lo} and {hi}.")
            ws.add_data_validation(dv)
            dv.add(f"{CL(left + j)}{hdr + 1}:{CL(left + j)}{last + 300}")
    return hdr + 1, last


def col_index(cols, h):
    for j, c in enumerate(cols):
        if c["h"] == h:
            return j
    raise KeyError(h)


# ======================================================================= DESIGN BASIS
DB_ITEMS = [
    # (key, label, unit, default, note, dv)  ; key None -> section header
    (None, "1. PROJECT IDENTIFICATION"),
    ("Project", "Project name", "text", None, "Free text", None),
    ("Location", "Site / location", "text", None, "Free text", None),
    ("DocNo", "Calculation document no.", "text", None, "", None),
    ("Rev", "Revision", "text", None, "", None),
    ("Author", "Prepared by", "text", None, "", None),
    ("Checker", "Checked by", "text", None, "Independent checker", None),
    ("Jurisdiction", "Governing code framework", "list", None,
     "Drives which standards rows apply in the Standards Matrix; does not change formulas", "=lstJuris"),
    (None, "2. BESS CONFIGURATION"),
    ("BESS_MW", "BESS rated power at POI", "MW", None, "Used for KPI normalisation only", None),
    ("BESS_MWh", "BESS rated energy", "MWh", None, "Used for KPI normalisation only", None),
    ("Tech", "BESS technology / chemistry", "text", None, "e.g. Li-ion LFP, NMC, flow, Na-ion (informative)", None),
    ("N_Encl", "Number of battery enclosures / containers", "no.", None,
     "May be referenced in Load Schedule quantity cells, e.g. =DB_N_Encl", None),
    ("N_PCS", "Number of PCS / inverter units", "no.", None, "May be referenced in Load Schedule quantity cells", None),
    ("N_MVS", "Number of MV skids / stations", "no.", None, "May be referenced in Load Schedule quantity cells", None),
    (None, "3. ELECTRICAL SYSTEM"),
    ("Freq", "System frequency", "Hz", None, "50 or 60 Hz. Equipment ratings entered must correspond to this frequency", "=lstFreq"),
    ("MV", "Default MV (transformer HV) nominal voltage", "kV", None, "Default where transformer row HV kV is blank", None),
    ("VLL", "Default LV system voltage, line-to-line", "V", None, "Default where transformer row LV is blank (e.g. 400, 415, 480, 600, 690)", None),
    ("VLN_ovr", "LV line-to-neutral voltage override", "V", None,
     "Leave blank for VLL/sqrt(3) (3-phase 4-wire). Enter for other systems", None),
    ("VLN", "LV line-to-neutral voltage (effective)", "V", "=IF(ISNUMBER(DB_VLN_ovr),DB_VLN_ovr,IF(ISNUMBER(DB_VLL),DB_VLL/SQRT3,\"\"))",
     "CALCULATED", None),
    ("VDC", "DC control / auxiliary system voltage", "V DC", None, "Used for DC-phase rows (informative DC load list)", None),
    ("Earthing", "LV system earthing arrangement", "list", None, "Informative; affects protection design, not load calc", "=lstEarth"),
    ("Arrangement", "Auxiliary transformer arrangement", "list", None, "Centralized or Distributed (transformers defined on sheet 6)", "=lstArr"),
    (None, "4. ENVIRONMENTAL / SITE"),
    ("Tamb", "Maximum design ambient temperature", "°C", None, "Site data (project specific)", None),
    ("Tavg", "Average ambient (24 h) at design condition", "°C", None, "Informative; relevant to loading-guide (thermal) assessments", None),
    ("Alt", "Site altitude above sea level", "m", None, "Site data", None),
    (None, "5. TRANSFORMER RATING BASIS & DERATING (verify against applicable standard / manufacturer)"),
    ("RefT_L", "Liquid-filled: rated max ambient (normal service condition)", "°C", 40,
     "Normal-service max ambient commonly 40 °C in IEC 60076-1 / IEEE C57.12.00 - VERIFY for edition & spec", None),
    ("RefAlt_L", "Liquid-filled: rated altitude limit", "m", 1000, "Commonly 1000 m - VERIFY", None),
    ("DerK_L", "Liquid-filled: kVA derating per K above rated ambient", "%/K", None,
     "NOT PROVIDED - obtain from manufacturer / loading guide. Blank = no ambient derating applied & flagged", None),
    ("DerAlt_L", "Liquid-filled: kVA derating per 100 m above altitude limit", "%/100 m", None,
     "NOT PROVIDED - obtain from applicable standard table / manufacturer. Blank = flagged", None),
    ("RefT_D", "Dry-type: rated max ambient", "°C", 40, "Commonly 40 °C (IEC 60076-11) - VERIFY", None),
    ("RefAlt_D", "Dry-type: rated altitude limit", "m", 1000, "Commonly 1000 m - VERIFY", None),
    ("DerK_D", "Dry-type: kVA derating per K above rated ambient", "%/K", None, "NOT PROVIDED - manufacturer", None),
    ("DerAlt_D", "Dry-type: kVA derating per 100 m above altitude limit", "%/100 m", None, "NOT PROVIDED - standard/manufacturer", None),
    ("StdSeries", "Standard rating series for recommendation", "list", "IEC R10",
     "Lists sheet. Recommended rating is NOT an approved selection", "=lstSeries"),
    (None, "6. DESIGN MARGINS & SIZING CRITERIA"),
    ("Margin", "Design margin on coincident demand", "%", None, "Project / client requirement. Blank = 0 % and flagged", None),
    ("Future", "Future expansion allowance", "%", None, "Project / client requirement. Blank = 0 % and flagged", None),
    ("DesignScen", "Default transformer design scenario", "list", "ENVELOPE",
     "ENVELOPE = highest of all scenarios per transformer", "=lstDesignScen"),
    ("MaxLoad", "Max permissible loading of rating (normal)", "%", 1.0, "100 % = rated (nameplate). Set per client spec", None),
    ("MaxLoadCont", "Max permissible loading under contingency (N-1)", "%", 1.0,
     ">100 % requires loading-guide study (IEC 60076-7 / IEEE C57.91 / C57.96)", None),
    ("RegLim", "Max transformer voltage regulation at design load", "%", None, "Project criterion (blank = REVIEW)", None),
    ("DipLim", "Max voltage dip during largest motor start", "%", None, "Project / equipment criterion (blank = REVIEW)", None),
    ("VDLim", "Max LV feeder voltage drop (steady state)", "%", None, "Project / code criterion (blank = REVIEW)", None),
    ("NLThr", "Non-linear load share triggering harmonic assessment", "%", None,
     "Engineering criterion (blank = REVIEW). Assessment per IEEE C57.110 / manufacturer", None),
    ("ImbLim", "Max LV phase-load imbalance", "%", None, "Project criterion (blank = REVIEW)", None),
    ("SrcMVA", "MV source short-circuit level at transformer HV", "MVA", None,
     "Blank = infinite source (conservative for fault level, optimistic for voltage dip) - flagged", None),
    ("Cmax", "Voltage factor c (max) for LV short-circuit", "-", None,
     "Per IEC 60909-0 table for LV systems (depends on voltage tolerance) - enter from standard", None),
    ("ContMult", "OCPD sizing multiplier on design current", "-", 1.0,
     "1.0 typical IEC practice; NEC continuous loads commonly 1.25 - VERIFY with code", None),
    ("K145", "Overload coordination factor for I2 check", "-", 1.45,
     "I2 <= 1.45 Iz per IEC 60364-4-43 (clause 433.1 - verify in purchased edition)", None),
    (None, "7. DEFAULTS APPLIED WHEN A LOAD-ROW FACTOR IS BLANK (visible, flagged as INFO in Load Schedule)"),
    ("DefDF", "Default demand (load) factor", "-", 1.0, "1.0 = conservative (full rated input)", None),
    ("DefDuty", "Default duty-cycle factor for energy", "-", 1.0, "1.0 = conservative", None),
    ("DefHours", "Default operating hours per day", "h/day", 24, "24 = conservative", None),
    (None, "8. REVIEW STATUS"),
    ("ProfRev", "Scenario profiles (sheet 4) reviewed for this project?", "list", "No", "Validation register item", "=lstYN"),
    ("StdRev", "Standards editions / clauses verified by responsible engineer?", "list", "No", "Validation register item", "=lstYN"),
]
PCT_KEYS = {"Margin", "Future", "MaxLoad", "MaxLoadCont", "RegLim", "DipLim", "VDLim", "NLThr", "ImbLim"}


def build_design_basis(wb, ws, cfg):
    title(ws, "Sheet 1 - Project Design Basis",
          "Yellow cells = user input (blue text). Grey = calculated. Each input is a workbook-level Name (DB_<key>) used throughout the model.")
    hdrs = ["Ref", "Parameter", "Value", "Unit", "Notes / source / status", "Input check", "Name"]
    for j, h in enumerate(hdrs):
        c = ws.cell(row=4, column=j + 1, value=h)
        c.fill = F_HDR
        c.font = FN_HDR
    for col, w in zip("ABCDEFG", [6, 52, 18, 10, 70, 30, 14]):
        ws.column_dimensions[col].width = w
    r = 5
    n = 0
    vals = cfg.get("db", {})
    for it in DB_ITEMS:
        if it[0] is None:
            ws.cell(row=r, column=1, value=it[1]).font = FN_B
            for j in range(1, 8):
                ws.cell(row=r, column=j).fill = F_SUB
            r += 1
            continue
        key, label, unit, default, note, dv = it
        n += 1
        ws.cell(row=r, column=1, value=f"DB{n:02d}")
        ws.cell(row=r, column=2, value=label)
        c = ws.cell(row=r, column=3)
        v = vals.get(key, default)
        is_calc = isinstance(default, str) and default.startswith("=")
        if is_calc:
            c.value = default
            c.fill = F_CALC
        else:
            if v is not None:
                c.value = v
            c.fill = F_IN
            c.font = FN_IN
        if key in PCT_KEYS:
            c.number_format = "0.0%"
            unit = "%"
        ws.cell(row=r, column=4, value=unit)
        ws.cell(row=r, column=5, value=note).alignment = WRAP
        ws.cell(row=r, column=7, value=f"DB_{key}")
        for j in range(1, 8):
            ws.cell(row=r, column=j).border = BRD
        add_name(wb, f"DB_{key}", f"'Design Basis'!$C${r}")
        if dv:
            d = DataValidation(type="list", formula1=dv, allow_blank=True, showErrorMessage=True,
                               error="Select from list")
            ws.add_data_validation(d)
            d.add(f"C{r}")
        # input checks
        chk = None
        required = {"Freq", "MV", "VLL", "Tamb", "Alt", "Arrangement", "Jurisdiction"}
        if key in required:
            chk = f'=IF(C{r}="","ERR: required input missing","OK")'
        if key == "Freq":
            chk = f'=IF(C{r}="","ERR: required input missing",IF(OR(C{r}=50,C{r}=60),"OK","ERR: 50 or 60 Hz only"))'
        if key in ("MV", "VLL", "VDC", "VLN_ovr"):
            msg = '"ERR: required input missing"' if key in required else '"OK (blank / not used)"'
            base = f'IF(C{r}="",{msg},'
            chk = f'={base}IF(AND(ISNUMBER(C{r}),C{r}>0),"OK","ERR: must be a number > 0"))'
        if key in ("Margin", "Future"):
            chk = f'=IF(C{r}="","REVIEW: blank - 0 % applied",IF(AND(ISNUMBER(C{r}),C{r}>=0,C{r}<=1),"OK","ERR: 0-100 %"))'
        if key in ("DerK_L", "DerAlt_L", "DerK_D", "DerAlt_D", "RegLim", "DipLim", "VDLim", "NLThr", "ImbLim",
                   "SrcMVA", "Cmax"):
            chk = f'=IF(C{r}="","REVIEW: not entered",IF(AND(ISNUMBER(C{r}),C{r}>=0),"OK","ERR: must be a number >= 0"))'
        if key in ("DefDF", "DefDuty"):
            chk = f'=IF(AND(ISNUMBER(C{r}),C{r}>=0,C{r}<=1),"OK","ERR: 0 to 1 required")'
        if key == "DefHours":
            chk = f'=IF(AND(ISNUMBER(C{r}),C{r}>=0,C{r}<=24),"OK","ERR: 0 to 24 h")'
        if key in ("MaxLoad", "MaxLoadCont"):
            chk = f'=IF(AND(ISNUMBER(C{r}),C{r}>0),"OK","ERR: > 0 required")'
        if key in ("ProfRev", "StdRev"):
            chk = f'=IF(C{r}="Yes","OK","REVIEW: not yet confirmed")'
        if chk:
            ws.cell(row=r, column=6, value=chk)
            ws.cell(row=r, column=6).fill = F_CALC
        r += 1
    status_cf(ws, f"F5:F{r}")
    add_name(wb, "DB_CheckRange", f"'Design Basis'!$F$5:$F${r}")
    ws.cell(row=r + 1, column=2, value="Derived").font = FN_B
    ws.cell(row=r + 2, column=2, value="BESS duration")
    ws.cell(row=r + 2, column=3, value='=IF(AND(ISNUMBER(DB_BESS_MW),DB_BESS_MW>0,ISNUMBER(DB_BESS_MWh)),DB_BESS_MWh/DB_BESS_MW,"")')
    ws.cell(row=r + 2, column=4, value="h")
    ws.cell(row=r + 2, column=3).fill = F_CALC
    ws.cell(row=r + 3, column=2, value="Number of auxiliary transformers defined (sheet 6)")
    ws.cell(row=r + 3, column=3, value='=COUNTIF(tblTX[TX_ID],"?*")')
    ws.cell(row=r + 3, column=3).fill = F_CALC
    ws.cell(row=r + 4, column=2, value="Design-basis errors (count)")
    ws.cell(row=r + 4, column=3, value=f'=COUNTIF(F5:F{r},"ERR*")')
    ws.cell(row=r + 4, column=3).fill = F_CALC
    add_name(wb, "DB_ErrCount", f"'Design Basis'!$C${r + 4}")
    ws.cell(row=r + 5, column=2, value="Design-basis review items (count)")
    ws.cell(row=r + 5, column=3, value=f'=COUNTIF(F5:F{r},"REVIEW*")')
    ws.cell(row=r + 5, column=3).fill = F_CALC
    add_name(wb, "DB_RevCount", f"'Design Basis'!$C${r + 5}")
    ws.freeze_panes = "A5"


# ======================================================================= LISTS
IEC_R10 = [25, 40, 50, 63, 80, 100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600, 2000,
           2500, 3150, 4000, 5000, 6300, 8000, 10000, 12500, 16000]
IEEE_3PH = [15, 30, 45, 75, 112.5, 150, 225, 300, 500, 750, 1000, 1500, 2000, 2500, 3750, 5000, 7500, 10000]


def build_lists(wb, ws, cfg):
    title(ws, "Lists - drop-down sources and standard rating series",
          "Editable. Rating series must be verified against the applicable standard / manufacturer offer. Add values in blank cells; keep each list contiguous.")
    lists = {
        "lstPhase": ["3PH", "1PH", "DC"],
        "lstConn": ["L-N", "L-L"],
        "lstPhAssign": ["L1", "L2", "L3", "L1-L2", "L2-L3", "L3-L1"],
        "lstRType": ["kW_in", "kVA", "A", "kW_out"],
        "lstYN": ["Yes", "No"],
        "lstNL": ["Y", "N"],
        "lstLagLead": ["Lag", "Lead"],
        "lstTXType": ["Liquid-filled", "Dry-type"],
        "lstCoolLiq": ["ONAN", "ONAF", "ONAN/ONAF", "KNAN", "KNAF", "KNAN/KNAF", "OFAF", "ODAF", "Other"],
        "lstCoolDry": ["AN", "AF", "AN/AF", "Other"],
        "lstScheme": ["N", "N+1", "2N"],
        "lstArr": ["Centralized", "Distributed"],
        "lstRole": ["Centralized", "Distributed section", "Dedicated (e.g. building)", "Standby / emergency"],
        "lstFreq": [50, 60],
        "lstJuris": ["IEC", "US (NEC/NFPA/IEEE)", "Mixed / Other"],
        "lstEarth": ["TN-S", "TN-C-S", "TT", "IT", "Solidly grounded wye", "Resistance grounded", "Ungrounded", "Other"],
        "lstScenCodes": SC,
        "lstDesignScen": ["ENVELOPE"] + SC,
        "lstSupply": ["Normal", "Essential", "UPS-backed", "Emergency (life safety)", "Optional standby"],
        "lstResult": ["PASS", "FAIL", "REVIEW", "NOT APPLICABLE"],
        "lstSeries": ["IEC R10", "IEEE 3-phase", "User list"],
    }
    col = 1
    for name, vals in lists.items():
        ws.cell(row=4, column=col, value=name).font = FN_B
        ws.cell(row=4, column=col).fill = F_SUB
        for i, v in enumerate(vals):
            c = ws.cell(row=5 + i, column=col, value=v)
            c.fill = F_IN
            c.font = FN_IN
        add_name(wb, name, f"Lists!${CL(col)}$5:${CL(col)}${4 + len(vals)}")
        ws.column_dimensions[CL(col)].width = max(12, len(name) + 2)
        col += 1
    # rating series
    col += 1
    hdr = ["k", "IEC R10 (kVA)", "IEEE 3-ph (kVA)", "User list (kVA)", "Active series (kVA)"]
    for j, h in enumerate(hdr):
        c = ws.cell(row=4, column=col + j, value=h)
        c.font = FN_B
        c.fill = F_SUB
        ws.column_dimensions[CL(col + j)].width = 16
    nrow = 40
    user = cfg.get("user_std", [])
    for i in range(nrow):
        r = 5 + i
        ws.cell(row=r, column=col, value=i + 1)
        for j, lst in enumerate([IEC_R10, IEEE_3PH, user]):
            c = ws.cell(row=r, column=col + 1 + j)
            if i < len(lst):
                c.value = lst[i]
            c.fill = F_IN
            c.font = FN_IN
        a, b, cc = (f"{CL(col + 1)}{r}", f"{CL(col + 2)}{r}", f"{CL(col + 3)}{r}")
        f = (f'=IF(DB_StdSeries="IEEE 3-phase",IF(ISNUMBER({b}),{b},""),IF(DB_StdSeries="User list",'
             f'IF(ISNUMBER({cc}),{cc},""),IF(ISNUMBER({a}),{a},"")))')
        c = ws.cell(row=r, column=col + 4, value=f)
        c.fill = F_CALC
    add_name(wb, "lstStdActive", f"Lists!${CL(col + 4)}$5:${CL(col + 4)}${4 + nrow}")
    ws.cell(row=46, column=col, value=(
        "IEC R10 = ISO 3 R10 preferred numbers as commonly used for IEC distribution transformers; IEEE list = three-phase "
        "ratings commonly listed in IEEE C57.12.00. Both are provided for convenience and MUST be verified against the "
        "current standard edition and manufacturer offer.")).font = FN_NOTE
    ws.freeze_panes = "A5"


# ======================================================================= LIBRARY
LIB_COLS = [
    dict(h="LibID", kind="in", w=12, unit="Unique ID"),
    dict(h="Category", kind="in", w=24, unit="from tblCat", dv="=lstCat"),
    dict(h="Description", kind="in", w=46, unit="Generic equipment description"),
    dict(h="Phase", kind="in", w=7, unit="3PH/1PH/DC", dv="=lstPhase"),
    dict(h="RatingType", kind="in", w=10, unit="kW_in / kVA / A / kW_out", dv="=lstRType"),
    dict(h="RatedValue", kind="in", w=10, fmt="0.00", unit="per unit, in units of RatingType (ENTER FROM DATASHEET)", dvnum=(0, 1e7)),
    dict(h="PF", kind="in", w=7, fmt="0.00", unit="0<PF<=1 (datasheet)", dvnum=(0.01, 1)),
    dict(h="Eff", kind="in", w=7, fmt="0.00", unit="0<η<=1 (kW_out only)", dvnum=(0.01, 1)),
    dict(h="StartMult", kind="in", w=9, fmt="0.0", unit="Starting kVA / rated kVA", dvnum=(1, 50)),
    dict(h="StartPF", kind="in", w=8, fmt="0.00", unit="PF during start", dvnum=(0.01, 1)),
    dict(h="Nonlinear", kind="in", w=9, unit="Y/N", dv="=lstNL"),
    dict(h="Profile", kind="in", w=12, unit="Default scenario profile", dv="=lstProf"),
    dict(h="Source", kind="in", w=28, unit="Datasheet / vendor doc reference"),
    dict(h="Notes", kind="in", w=50, unit="Engineering notes"),
    dict(h="Status", kind="calc", w=26, unit="Library entry check",
         f='IF([@LibID]="","",IF(COUNTIF(tblLib[LibID],[@LibID])>1,"ERR: duplicate LibID",'
           'IF(ISNUMBER([@RatedValue]),IF([@Source]="","WARN: rating has no source","OK - rated"),'
           '"INFO: rating not entered (enter in library or load row)")))'),
]

CATS = [
    ("BESS-ENC", "BESS container / enclosure auxiliaries (BMS, lighting, controls, internal fans)"),
    ("BESS-THERM", "BESS HVAC and liquid cooling (chillers, pumps, heaters)"),
    ("PCS-AUX", "PCS / inverter auxiliaries (fans, cooling, control power, heaters)"),
    ("MV-SWGR", "MV switchgear & breaker auxiliaries (motors, heaters, protection relays)"),
    ("TX-AUX", "Transformer cooling & control auxiliaries (fans, pumps, OLTC, monitoring)"),
    ("DC-SYS", "Battery chargers & DC systems"),
    ("UPS", "UPS systems"),
    ("CTRL-COMM", "SCADA, EMS, BMS master, PPC & telecommunications"),
    ("FIRE", "Fire detection, gas detection & suppression systems"),
    ("FIRE-PUMP", "Fire pumps & fire-water systems (where applicable)"),
    ("LIGHT", "Lighting (normal, emergency, area / perimeter)"),
    ("RECEPT", "Receptacles / small power / maintenance sockets"),
    ("SECURITY", "Security, CCTV, access control, intrusion detection"),
    ("BLDG", "Control building & O&M facilities (HVAC, domestic, workshop)"),
    ("SUBSTN", "MV yard / HV substation auxiliaries"),
    ("MISC", "Miscellaneous / user-defined"),
]

LIB_DEFAULT = [
    # LibID, cat, description, phase, rtype, nonlinear, profile, notes
    ("ENC-BMS", "BESS-ENC", "Battery management system (rack/bank BMS controllers) per enclosure", "1PH", "kW_in", "Y", "CTRL", "Typically UPS-backed"),
    ("ENC-CTRL", "BESS-ENC", "Enclosure control panel, I/O, network switches", "1PH", "kW_in", "Y", "CTRL", ""),
    ("ENC-LIGHT", "BESS-ENC", "Enclosure internal lighting", "1PH", "kW_in", "Y", "LIGHT", ""),
    ("ENC-VENT", "BESS-ENC", "Enclosure ventilation / exhaust fan (incl. deflagration-related ventilation)", "3PH", "kW_out", "N", "CONT", "Check NFPA 855 / UL 9540A-derived ventilation needs"),
    ("TH-HVAC", "BESS-THERM", "Enclosure HVAC unit (air-cooled) - compressor + fans", "3PH", "kW_in", "Y", "COOL", "Rate at electrical input at design ambient; inverter-driven = non-linear"),
    ("TH-CHILL", "BESS-THERM", "Liquid-cooling chiller / thermal management unit", "3PH", "kW_in", "Y", "COOL", "Electrical input at max design ambient"),
    ("TH-PUMP", "BESS-THERM", "Coolant circulation pump", "3PH", "kW_out", "N", "COOL", "Shaft rating -> use efficiency"),
    ("TH-HEAT", "BESS-THERM", "Enclosure / coolant heater", "3PH", "kW_in", "N", "HEAT", "Use exclusive group with cooling where interlocked"),
    ("PCS-FAN", "PCS-AUX", "PCS cooling fans / cooling unit", "3PH", "kW_in", "Y", "PCS", ""),
    ("PCS-CTRL", "PCS-AUX", "PCS control power supply", "1PH", "kW_in", "Y", "CTRL", ""),
    ("PCS-HEAT", "PCS-AUX", "PCS anti-condensation heaters", "1PH", "kW_in", "N", "HEAT", ""),
    ("MV-MOTOR", "MV-SWGR", "MV breaker spring-charging motors", "1PH", "kW_in", "N", "INTERMIT", "Intermittent; inrush relevant"),
    ("MV-HEAT", "MV-SWGR", "MV switchgear anti-condensation heaters", "1PH", "kW_in", "N", "CONT", ""),
    ("MV-PROT", "MV-SWGR", "Protection relays & IEDs (usually DC-supplied)", "DC", "kW_in", "Y", "CTRL", "DC row - informative"),
    ("TX-FAN", "TX-AUX", "Transformer cooling fans (ONAF stage)", "3PH", "kW_out", "N", "TXCOOL", ""),
    ("TX-PUMP", "TX-AUX", "Transformer oil pumps (OF/OD cooling)", "3PH", "kW_out", "N", "TXCOOL", ""),
    ("TX-OLTC", "TX-AUX", "OLTC motor drive", "3PH", "kW_in", "N", "INTERMIT", ""),
    ("TX-MON", "TX-AUX", "Transformer monitoring / control cabinet & heaters", "1PH", "kW_in", "N", "CONT", ""),
    ("DC-CHG", "DC-SYS", "Battery charger / rectifier (AC input)", "3PH", "kVA", "Y", "CHARGER", "Size for DC load + battery recharge"),
    ("UPS-01", "UPS", "UPS (AC input incl. battery recharge)", "3PH", "kVA", "Y", "CHARGER", "Rectifier input rating incl. recharge"),
    ("CC-SCADA", "CTRL-COMM", "SCADA / EMS / PPC servers & HMI", "1PH", "kW_in", "Y", "CTRL", ""),
    ("CC-TEL", "CTRL-COMM", "Telecommunications, fibre & network equipment", "1PH", "kW_in", "Y", "CTRL", ""),
    ("FI-DET", "FIRE", "Fire / gas detection & alarm panel", "1PH", "kW_in", "Y", "CTRL", "Secondary supply rules per fire code"),
    ("FI-SUPP", "FIRE", "Suppression system control / release panel", "1PH", "kW_in", "Y", "CTRL", ""),
    ("FP-ELEC", "FIRE-PUMP", "Electric fire pump motor", "3PH", "kW_out", "N", "FIREPUMP", "Supply arrangements per NFPA 20 / local code if applicable"),
    ("FP-JOCK", "FIRE-PUMP", "Jockey pump", "3PH", "kW_out", "N", "CONT", ""),
    ("LT-AREA", "LIGHT", "Area / perimeter lighting", "1PH", "kW_in", "Y", "LIGHTNIGHT", ""),
    ("LT-BLDG", "LIGHT", "Building internal lighting", "1PH", "kW_in", "Y", "LIGHT", ""),
    ("LT-EMER", "LIGHT", "Emergency / egress lighting", "1PH", "kW_in", "Y", "CONT", ""),
    ("RC-GEN", "RECEPT", "General receptacles / small power", "1PH", "kVA", "N", "MAINTONLY", ""),
    ("RC-MAINT", "RECEPT", "Maintenance power outlets (3-phase)", "3PH", "kVA", "N", "MAINTONLY", ""),
    ("SE-CCTV", "SECURITY", "CCTV cameras, NVR & PoE switches", "1PH", "kW_in", "Y", "CTRL", ""),
    ("SE-ACC", "SECURITY", "Access control & intrusion detection", "1PH", "kW_in", "Y", "CTRL", ""),
    ("SE-GATE", "SECURITY", "Motorised gate", "1PH", "kW_out", "N", "INTERMIT", ""),
    ("BL-HVAC", "BLDG", "Control / O&M building HVAC", "3PH", "kW_in", "Y", "COOL", ""),
    ("BL-HEAT", "BLDG", "Building space heating", "3PH", "kW_in", "N", "HEAT", ""),
    ("BL-WH", "BLDG", "Water heater / domestic loads", "1PH", "kW_in", "N", "CONT", ""),
    ("BL-WSHOP", "BLDG", "Workshop equipment", "3PH", "kVA", "N", "MAINTONLY", ""),
    ("SS-AUX", "SUBSTN", "HV substation AC auxiliaries (marshalling, heaters, disconnector motors)", "3PH", "kVA", "N", "CONT", ""),
    ("SS-LIGHT", "SUBSTN", "Substation yard lighting", "1PH", "kW_in", "Y", "LIGHTNIGHT", ""),
    ("MI-USER", "MISC", "User-defined load", "3PH", "kW_in", "N", "CONT", ""),
]


def build_library(wb, ws, cfg):
    title(ws, "Sheet 2 - Equipment Library",
          "Generic, reusable equipment classes. RATINGS ARE INTENTIONALLY BLANK - enter from project datasheets with source reference. "
          "Add rows at the bottom of either table (tables expand automatically).")
    cats = cfg.get("cats", CATS)
    ws.cell(row=3, column=1, value="Load categories (tblCat)").font = FN_B
    ccols = [dict(h="Category", kind="in", w=14, unit="Code"), dict(h="CatDescription", kind="in", w=40, unit="Description")]
    make_table(wb, ws, "tblCat", ccols, [{"Category": a, "CatDescription": b} for a, b in cats], top=5, left=1, nblank=4)
    add_name(wb, "lstCat", "tblCat[Category]")
    lib = cfg.get("lib")
    if lib is None:
        lib = [dict(LibID=a, Category=b, Description=c, Phase=d, RatingType=e, Nonlinear=f, Profile=g, Notes=h)
               for a, b, c, d, e, f, g, h in LIB_DEFAULT]
    ws.cell(row=3, column=4, value="Equipment library (tblLib)").font = FN_B
    make_table(wb, ws, "tblLib", LIB_COLS, lib, top=5, left=4, nblank=10)
    add_name(wb, "lstLib", "tblLib[LibID]")
    status_cf(ws, f"{CL(4 + len(LIB_COLS) - 1)}6:{CL(4 + len(LIB_COLS) - 1)}400")
    ws.freeze_panes = "E6"


# ======================================================================= PROFILES / SCENARIOS
PROFILES = [
    ("CONT", "Continuous in all scenarios", [1, 1, 1, 1, 1, 1, 1]),
    ("CTRL", "Control / protection / comms / fire detection (essential)", [1, 1, 1, 1, 1, 1, 1]),
    ("COOL", "Cooling (HVAC / chiller) - set NORM to expected fraction", [1, 1, 1, 1, 1, 0, 0]),
    ("HEAT", "Heating (exclusive with cooling via ExclGroup)", [1, 0, 1, 1, 1, 0, 0]),
    ("PCS", "PCS auxiliaries in operation", [1, 1, 1, 1, 0, 0, 1]),
    ("TXCOOL", "Transformer forced-cooling stage", [1, 1, 1, 0, 0, 0, 0]),
    ("CHARGER", "Chargers / UPS input incl. recharge", [1, 1, 1, 1, 1, 1, 1]),
    ("LIGHT", "Normal lighting", [1, 1, 1, 1, 1, 0, 1]),
    ("LIGHTNIGHT", "Area lighting (night-time)", [1, 1, 1, 1, 1, 0, 1]),
    ("MAINTONLY", "Maintenance-only loads", [0, 0, 1, 0, 1, 0, 0]),
    ("INTERMIT", "Intermittent motor loads (count in max-sim & maintenance)", [0, 0, 1, 0, 1, 0, 1]),
    ("FIREPUMP", "Fire pump - emergency only", [0, 0, 0, 0, 0, 1, 0]),
    ("ESSENTIAL", "Essential loads on emergency supply only", [0, 0, 0, 0, 0, 1, 1]),
    ("OFF", "Installed but not operating (spare)", [0, 0, 0, 0, 0, 0, 0]),
]


def build_scenarios(wb, ws, cfg):
    title(ws, "Sheet 4 - Operating Scenario Analysis",
          "Scenario participation factors (0-1) per profile. Mutually exclusive loads (same ExclGroup in Load Schedule) are "
          "counted once per scenario: the highest-kW member only. Defaults are ENGINEERING ASSUMPTIONS - review for the project.")
    ws.cell(row=3, column=1, value="Scenario definitions (tblScen)").font = FN_B
    scols = [dict(h="Code", kind="in", w=11), dict(h="ScenarioName", kind="in", w=32), dict(h="Definition", kind="in", w=60)]
    defs = ["BESS charging/discharging at rated conditions, typical ambient",
            "BESS at rated throughput with maximum design ambient (all cooling at max)",
            "All loads that can coincide physically, incl. intermittent & maintenance loads",
            "BESS energised, no power exchange; thermal management for SOC hold",
            "Planned maintenance: maintenance power, workshop, lighting; PCS off",
            "Loss of normal supply / fire event: essential + life-safety + fire pumps",
            "Restoration from dead plant: control, PCS aux and charging only"]
    make_table(wb, ws, "tblScen", scols, [dict(Code=c, ScenarioName=n, Definition=d) for c, n, d in zip(SC, SC_NAME, defs)],
               top=5, left=1)
    ws.cell(row=14, column=1, value="Scenario participation profiles (tblProf)").font = FN_B
    pcols = [dict(h="Profile", kind="in", w=11), dict(h="ProfDescription", kind="in", w=40)] + \
            [dict(h=s, kind="in", w=9, fmt="0.00", dvnum=(0, 1)) for s in SC]
    profs = cfg.get("profiles", PROFILES)
    rows = []
    for p, d, v in profs:
        row = {"Profile": p, "ProfDescription": d}
        row.update(dict(zip(SC, v)))
        rows.append(row)
    f0, l0 = make_table(wb, ws, "tblProf", pcols, rows, top=16, left=1, nblank=4)
    add_name(wb, "lstProf", "tblProf[Profile]")
    # results
    r0 = l0 + 3
    ws.cell(row=r0, column=1, value="Site scenario results (all AC loads, all transformers)").font = FN_B
    hd = ["Code", "Scenario", "ΣP raw (kW)", "ΣP counted (kW)", "Excluded by mutual exclusivity (kW)", "ΣQ (kvar)",
          "S vector (kVA)", "S arithmetic (kVA)", "PF", "Largest start Δ (kVA)", "Transient S (kVA)",
          "Rows contributing", "Exclusive groups resolved", "Rank (by S)"]
    for j, h in enumerate(hd):
        c = ws.cell(row=r0 + 1, column=1 + j, value=h)
        c.fill = F_HDR
        c.font = FN_HDR
        c.alignment = WRAP
    for i, s in enumerate(SC):
        r = r0 + 2 + i
        f = {
            1: f"=INDEX(tblScen[Code],{i + 1})",
            2: f"=INDEX(tblScen[ScenarioName],{i + 1})",
            3: f"=SUM(tblLoad[R_{s}])",
            4: f"=SUM(tblLoad[P_{s}])",
            5: f"=C{r}-D{r}",
            6: f"=SUMPRODUCT(tblLoad[P_{s}],tblLoad[TanPhi])",
            7: f"=SQRT(D{r}^2+F{r}^2)",
            8: f"=SUMPRODUCT(tblLoad[P_{s}]*SQRT(1+tblLoad[TanPhi]^2))",
            9: f'=IF(G{r}>0,D{r}/G{r},"")',
            10: f"=SUMPRODUCT(MAX((tblLoad[P_{s}]>0)*tblLoad[StartDelta]))",
            11: f"=G{r}+J{r}",
            12: f'=COUNTIF(tblLoad[P_{s}],">0")',
            13: f'=SUMPRODUCT((tblLoad[ExclGroup]<>"")*(tblLoad[R_{s}]>0)*(tblLoad[P_{s}]=0))',
            14: f"=RANK(G{r},$G${r0 + 2}:$G${r0 + 8})",
        }
        for j, fx in f.items():
            c = ws.cell(row=r, column=j, value=fx)
            c.fill = F_CALC
            c.border = BRD
            if j in (3, 4, 5, 6, 7, 8, 10, 11):
                c.number_format = "#,##0.0"
            if j == 9:
                c.number_format = "0.000"
    add_name(wb, "SCN_S", f"'Scenario Analysis'!$G${r0 + 2}:$G${r0 + 8}")
    add_name(wb, "SCN_P", f"'Scenario Analysis'!$D${r0 + 2}:$D${r0 + 8}")
    add_name(wb, "SCN_Names", f"'Scenario Analysis'!$B${r0 + 2}:$B${r0 + 8}")
    add_name(wb, "SCN_Codes", f"'Scenario Analysis'!$A${r0 + 2}:$A${r0 + 8}")
    add_name(wb, "SCN_Raw", f"'Scenario Analysis'!$C${r0 + 2}:$C${r0 + 8}")
    add_name(wb, "SCN_Trans", f"'Scenario Analysis'!$K${r0 + 2}:$K${r0 + 8}")
    rr = r0 + 10
    ws.cell(row=rr, column=1, value="Peak scenario (max S vector)").font = FN_B
    ws.cell(row=rr, column=3, value="=INDEX(SCN_Names,MATCH(MAX(SCN_S),SCN_S,0))").fill = F_CALC
    ws.cell(row=rr, column=7, value="=MAX(SCN_S)").fill = F_CALC
    ws.cell(row=rr + 1, column=1, value="Consistency: counted ≤ raw for every scenario").font = FN_B
    ws.cell(row=rr + 1, column=3, value=f'=IF(SUMPRODUCT((SCN_P>SCN_Raw+0.000001)*1)=0,"PASS","FAIL")').fill = F_CALC
    add_name(wb, "SCN_Consistency", f"'Scenario Analysis'!$C${rr + 1}")
    ws.cell(row=rr + 2, column=1, value="Check: Max-simultaneous ≥ Normal and ≥ Max-cooling").font = FN_B
    ws.cell(row=rr + 2, column=3, value=f'=IF(AND(G{r0 + 4}>=G{r0 + 2}-0.000001,G{r0 + 4}>=G{r0 + 3}-0.000001),"PASS","REVIEW: MAXSIM lower than NORM/MAXCOOL - check profiles")').fill = F_CALC
    add_name(wb, "SCN_MaxSimChk", f"'Scenario Analysis'!$C${rr + 2}")
    status_cf(ws, f"C{rr + 1}:C{rr + 2}")
    # TX x scenario matrix
    rm = rr + 5
    ws.cell(row=rm, column=1, value="Transformer × scenario apparent power (kVA) - first 20 transformers in tblTX").font = FN_B
    hd = ["#", "TX_ID"] + SC + ["Design kVA", "Governing"]
    for j, h in enumerate(hd):
        c = ws.cell(row=rm + 1, column=1 + j, value=h)
        c.fill = F_HDR
        c.font = FN_HDR
    for k in range(1, 21):
        r = rm + 1 + k
        ws.cell(row=r, column=1, value=k)
        ws.cell(row=r, column=2, value=f'=IF(ROWS(tblTX[TX_ID])>={k},IF(INDEX(tblTX[TX_ID],{k})="","",INDEX(tblTX[TX_ID],{k})),"")')
        for j, s in enumerate(SC + ["Design_kVA", "Gov"]):
            c = ws.cell(row=r, column=3 + j,
                        value=f'=IF($B{r}="","",INDEX(tblTX[{"S_" + s if s in SC else s}],{k}))')
            c.fill = F_CALC
            c.number_format = "#,##0.0"
        ws.cell(row=r, column=2).fill = F_CALC
    ws.freeze_panes = "A6"


# ======================================================================= LOAD SCHEDULE
def load_cols():
    cols = [
        dict(h="ID", kind="in", w=10, unit="Unique equipment tag"),
        dict(h="LibID", kind="in", w=11, unit="Optional: library item (fills blanks)", dv="=lstLib"),
        dict(h="Description", kind="in", w=38, unit="Equipment description"),
        dict(h="Category", kind="in", w=11, unit="Blank = library", dv="=lstCat"),
        dict(h="TX", kind="in", w=9, unit="Supply transformer / group (sheet 6)", dv="=lstTX"),
        dict(h="Feeder", kind="in", w=10, unit="LV feeder (sheet 7)", dv="=lstFdr"),
        dict(h="QtyInst", kind="in", w=7, fmt="0.##", unit="Qty installed", dvnum=(0, 100000)),
        dict(h="QtyDuty", kind="in", w=7, fmt="0.##", unit="Qty running (duty). Blank = installed", dvnum=(0, 100000)),
        dict(h="Phase", kind="in", w=6, unit="3PH/1PH/DC. Blank = library", dv="=lstPhase"),
        dict(h="Conn1Ph", kind="in", w=6, unit="1PH: L-N / L-L", dv="=lstConn"),
        dict(h="PhaseAssign", kind="in", w=7, unit="1PH phase allocation", dv="=lstPhAssign"),
        dict(h="VOverride", kind="in", w=8, fmt="0", unit="Voltage override (V). Blank = system", dvnum=(1, 100000)),
        dict(h="RatingType", kind="in", w=8, unit="kW_in/kVA/A/kW_out. Blank = library", dv="=lstRType"),
        dict(h="Rating", kind="in", w=9, fmt="0.000", unit="Rating per unit (units per RatingType). Blank = library", dvnum=(0, 1e7)),
        dict(h="PF", kind="in", w=6, fmt="0.00", unit="Running PF. Blank = library", dvnum=(0.01, 1)),
        dict(h="Reactive", kind="in", w=6, unit="Lag / Lead (blank = Lag)", dv="=lstLagLead"),
        dict(h="Eff", kind="in", w=6, fmt="0.00", unit="Efficiency (kW_out only)", dvnum=(0.01, 1)),
        dict(h="DF", kind="in", w=6, fmt="0.00", unit="Demand (load) factor 0-1. Blank = DB default", dvnum=(0, 1)),
        dict(h="DutyF", kind="in", w=6, fmt="0.00", unit="Duty-cycle factor (energy) 0-1", dvnum=(0, 1)),
        dict(h="Hours", kind="in", w=6, fmt="0.0", unit="Operating h/day", dvnum=(0, 24)),
        dict(h="Profile", kind="in", w=10, unit="Scenario profile. Blank = library", dv="=lstProf"),
        dict(h="ExclGroup", kind="in", w=9, unit="Mutual-exclusivity group (text)"),
        dict(h="SupplyClass", kind="in", w=11, unit="Normal / essential / UPS / emergency", dv="=lstSupply"),
        dict(h="StartMult", kind="in", w=7, fmt="0.0", unit="Start kVA ÷ rated kVA. Blank = library / 1", dvnum=(1, 50)),
        dict(h="StartPF", kind="in", w=6, fmt="0.00", unit="PF at start", dvnum=(0.01, 1)),
        dict(h="Nonlinear", kind="in", w=6, unit="Y/N. Blank = library", dv="=lstNL"),
        dict(h="THDi", kind="in", w=6, fmt="0%", unit="Current THD (info)", dvnum=(0, 5)),
        dict(h="Reference", kind="in", w=18, unit="Datasheet / document ref"),
        dict(h="Remarks", kind="in", w=24, unit="Engineering remarks"),
        # ---- calculated
        dict(h="Check", kind="calc", w=34, unit="Row validation (first issue)"),
        dict(h="Active", kind="calc", w=5, f='IF(AND([@ID]="",[@LibID]="",[@Description]="",[@Rating]=""),0,1)', unit="1 = row in use"),
        dict(h="Row", kind="calc", w=5, f="ROW()-ROW(tblLoad[[#Headers],[ID]])", unit="Row no."),
        dict(h="LibIdx", kind="calc", w=5, f='IF([@LibID]="",0,IFERROR(MATCH([@LibID],tblLib[LibID],0),-1))', unit="Library row (-1 = not found)"),
        dict(h="Cat_e", kind="calc", w=10, f='IF([@Category]<>"",[@Category],IF([@LibIdx]>0,INDEX(tblLib[Category],[@LibIdx])&"",""))', unit="Effective category"),
        dict(h="Phase_e", kind="calc", w=6, f='IF([@Phase]<>"",[@Phase],IF([@LibIdx]>0,INDEX(tblLib[Phase],[@LibIdx])&"",""))', unit="Effective phase"),
        dict(h="RType_e", kind="calc", w=7, f='IF([@RatingType]<>"",[@RatingType],IF([@LibIdx]>0,INDEX(tblLib[RatingType],[@LibIdx])&"",""))', unit="Effective rating type"),
    ]

    def num(inp, libcol):
        return (f'IF(ISNUMBER([@{inp}]),[@{inp}],IF([@LibIdx]>0,IF(ISNUMBER(INDEX(tblLib[{libcol}],[@LibIdx])),'
                f'INDEX(tblLib[{libcol}],[@LibIdx]),""),""))')
    cols += [
        dict(h="Rating_e", kind="calc", w=8, fmt="0.000", f=num("Rating", "RatedValue"), unit="Effective rating"),
        dict(h="PF_e", kind="calc", w=6, fmt="0.00", f='IF([@Phase_e]="DC",1,' + num("PF", "PF") + ")", unit="Effective PF (DC = 1)"),
        dict(h="Eff_e", kind="calc", w=6, fmt="0.00", f=num("Eff", "Eff"), unit="Effective efficiency"),
        dict(h="StartM_e", kind="calc", w=6, fmt="0.0", f=num("StartMult", "StartMult"), unit="Effective start multiplier"),
        dict(h="StartPF_e", kind="calc", w=6, fmt="0.00", f=num("StartPF", "StartPF"), unit="Effective start PF"),
        dict(h="NL_e", kind="calc", w=5, f='IF([@Nonlinear]<>"",[@Nonlinear],IF([@LibIdx]>0,INDEX(tblLib[Nonlinear],[@LibIdx])&"",""))', unit="Effective non-linear flag"),
        dict(h="Prof_e", kind="calc", w=9, f='IF([@Profile]<>"",[@Profile],IF([@LibIdx]>0,INDEX(tblLib[Profile],[@LibIdx])&"",""))', unit="Effective profile"),
        dict(h="QtyDuty_e", kind="calc", w=6, fmt="0.##", f='IF(ISNUMBER([@QtyDuty]),[@QtyDuty],N([@QtyInst]))', unit="Effective duty qty"),
        dict(h="DF_e", kind="calc", w=5, fmt="0.00", f='IF(ISNUMBER([@DF]),[@DF],DB_DefDF)', unit="Effective DF"),
        dict(h="DutyF_e", kind="calc", w=5, fmt="0.00", f='IF(ISNUMBER([@DutyF]),[@DutyF],DB_DefDuty)', unit="Effective duty factor"),
        dict(h="Hours_e", kind="calc", w=5, fmt="0.0", f='IF(ISNUMBER([@Hours]),[@Hours],DB_DefHours)', unit="Effective h/day"),
        dict(h="VLL", kind="calc", w=6, fmt="0", f='IF([@TX]="",N(DB_VLL),IFERROR(INDEX(tblTX[VLL_e],MATCH([@TX],tblTX[TX_ID],0)),N(DB_VLL)))', unit="System V L-L"),
        dict(h="VLN", kind="calc", w=6, fmt="0", f='IF([@TX]="",N(DB_VLN),IFERROR(INDEX(tblTX[VLN_e],MATCH([@TX],tblTX[TX_ID],0)),N(DB_VLN)))', unit="System V L-N"),
        dict(h="V_app", kind="calc", w=6, fmt="0", f='IF(ISNUMBER([@VOverride]),[@VOverride],IF([@Phase_e]="3PH",[@VLL],IF([@Phase_e]="1PH",IF([@Conn1Ph]="L-L",[@VLL],[@VLN]),IF([@Phase_e]="DC",N(DB_VDC),0))))', unit="Applied voltage (V)"),
        dict(h="PfIdx", kind="calc", w=5, f='IFERROR(MATCH([@Prof_e],tblProf[Profile],0),0)', unit="Profile row"),
        dict(h="Valid", kind="calc", w=5, f='IF(AND([@Active]=1,LEFT([@Check],3)<>"ERR"),1,0)', unit="1 = included in totals"),
        dict(h="S_unit", kind="calc", w=8, fmt="0.000",
             f='IF([@Valid]=0,0,IF([@RType_e]="kVA",[@Rating_e],IF([@RType_e]="A",IF([@Phase_e]="3PH",SQRT3*[@V_app]*[@Rating_e]/1000,[@V_app]*[@Rating_e]/1000),'
               'IF([@RType_e]="kW_in",[@Rating_e]/[@PF_e],IF([@RType_e]="kW_out",[@Rating_e]/[@Eff_e]/[@PF_e],0)))))',
             unit="kVA input per unit"),
        dict(h="P_unit", kind="calc", w=8, fmt="0.000", f='IF([@Valid]=0,0,[@S_unit]*[@PF_e])', unit="kW input per unit"),
        dict(h="I_unit", kind="calc", w=8, fmt="0.00", f='IF([@Valid]=0,0,IF([@Phase_e]="3PH",[@S_unit]*1000/(SQRT3*[@V_app]),[@S_unit]*1000/[@V_app]))', unit="A per unit (line)"),
        dict(h="TanPhi", kind="calc", w=6, fmt="0.000", f='IF([@Valid]=0,0,IF([@Reactive]="Lead",-1,1)*SQRT(1-[@PF_e]^2)/[@PF_e])', unit="tan φ (+ lagging / − leading)"),
        dict(h="IsAC", kind="calc", w=4, f='IF(AND([@Valid]=1,[@Phase_e]<>"DC"),1,0)', unit="AC row"),
        dict(h="Conn_kW", kind="calc", w=8, fmt="#,##0.00", f='[@IsAC]*N([@QtyInst])*[@P_unit]', unit="Connected kW (installed qty)"),
        dict(h="Conn_kVA", kind="calc", w=8, fmt="#,##0.00", f='[@IsAC]*N([@QtyInst])*[@S_unit]', unit="Connected kVA"),
        dict(h="Run_kW", kind="calc", w=8, fmt="#,##0.00", f='[@IsAC]*[@QtyDuty_e]*[@P_unit]*[@DF_e]', unit="Max demand kW (duty qty × DF)"),
        dict(h="Run_kVA", kind="calc", w=8, fmt="#,##0.00", f='[@IsAC]*[@QtyDuty_e]*[@S_unit]*[@DF_e]', unit="Max demand kVA"),
        dict(h="DC_kW", kind="calc", w=7, fmt="#,##0.00", f='IF(AND([@Valid]=1,[@Phase_e]="DC"),[@QtyDuty_e]*[@P_unit]*[@DF_e],0)', unit="DC load kW (not in AC totals)"),
        dict(h="Start_kVA", kind="calc", w=8, fmt="#,##0.00", f='[@IsAC]*[@S_unit]*IF(ISNUMBER([@StartM_e]),[@StartM_e],1)', unit="Starting kVA of ONE unit"),
        dict(h="Start_kW", kind="calc", w=8, fmt="#,##0.00", f='[@Start_kVA]*IF(ISNUMBER([@StartPF_e]),[@StartPF_e],N([@PF_e]))', unit="Starting kW of one unit"),
        dict(h="StartDelta", kind="calc", w=8, fmt="#,##0.00", f='MAX(0,[@Start_kVA]-[@IsAC]*[@S_unit])', unit="Start kVA above running"),
        dict(h="Day_kWh", kind="calc", w=9, fmt="#,##0.0", f='[@Run_kW]*[@DutyF_e]*[@Hours_e]', unit="Daily energy kWh"),
    ]
    for s in SC:
        cols.append(dict(h=f"R_{s}", kind="calc", w=8, fmt="#,##0.00",
                         f=f'IF([@PfIdx]=0,0,[@Run_kW]*INDEX(tblProf[{s}],[@PfIdx]))', unit=f"Raw kW {s}"))
    for s in SC:
        cols.append(dict(h=f"RS_{s}", kind="calc", w=8, fmt="#,##0.00",
                         f=f'[@R_{s}]*SQRT(1+[@TanPhi]^2)', unit=f"Raw kVA {s} (exclusivity metric)"))
    for s in SC:
        cols.append(dict(h=f"P_{s}", kind="calc", w=8, fmt="#,##0.00",
                         f=(f'IF([@ExclGroup]="",[@R_{s}],IF(AND([@RS_{s}]=SUMPRODUCT(MAX((tblLoad[ExclGroup]=[@ExclGroup])*tblLoad[RS_{s}])),'
                            f'SUMPRODUCT((tblLoad[ExclGroup]=[@ExclGroup])*(tblLoad[RS_{s}]=[@RS_{s}])*(tblLoad[Row]<[@Row]))=0),[@R_{s}],0))'),
                         unit=f"Counted kW {s} (exclusivity by kVA; own P/Q pair kept)"))
    chk = ('IF([@Active]=0,"",IF([@LibIdx]=-1,"ERR: LibID not in Equipment Library",'
           'IF([@Phase_e]="","ERR: phase missing",'
           'IF(AND([@Phase_e]<>"3PH",[@Phase_e]<>"1PH",[@Phase_e]<>"DC"),"ERR: invalid phase",'
           'IF(AND([@RType_e]<>"kW_in",[@RType_e]<>"kVA",[@RType_e]<>"A",[@RType_e]<>"kW_out"),"ERR: rating type missing/invalid",'
           'IF(NOT(ISNUMBER([@Rating_e])),"ERR: rating missing (row or library)",'
           'IF([@Rating_e]<0,"ERR: negative rating",'
           'IF(OR(NOT(ISNUMBER([@PF_e])),[@PF_e]<=0,[@PF_e]>1),"ERR: PF must satisfy 0 < PF <= 1",'
           'IF(AND([@RType_e]="kW_out",OR(NOT(ISNUMBER([@Eff_e])),[@Eff_e]<=0,[@Eff_e]>1)),"ERR: efficiency 0<η<=1 required for kW_out",'
           'IF(OR(NOT(ISNUMBER([@QtyInst])),[@QtyInst]<=0),"ERR: installed quantity must be > 0",'
           'IF(OR([@QtyDuty_e]<0,[@QtyDuty_e]>[@QtyInst]),"ERR: duty qty must be 0..installed qty",'
           'IF(OR([@DF_e]<0,[@DF_e]>1,[@DutyF_e]<0,[@DutyF_e]>1),"ERR: DF / duty factor outside 0-1",'
           'IF(OR([@Hours_e]<0,[@Hours_e]>24),"ERR: hours/day outside 0-24",'
           'IF([@PfIdx]=0,"ERR: scenario profile missing/invalid",'
           'IF(AND([@TX]<>"",ISNA(MATCH([@TX],tblTX[TX_ID],0))),"ERR: transformer not defined on sheet 6",'
           'IF(AND([@Feeder]<>"",ISNA(MATCH([@Feeder],tblFdr[FeederID],0))),"ERR: feeder not defined on sheet 7",'
           'IF(AND([@Feeder]<>"",[@TX]<>"",IFERROR(INDEX(tblFdr[TX_ID],MATCH([@Feeder],tblFdr[FeederID],0)),[@TX])<>[@TX]),"ERR: feeder is supplied from a different transformer",'
           'IF(OR(NOT(ISNUMBER([@V_app])),[@V_app]<=0),"ERR: voltage undefined (set LV/DC voltage or override)",'
           'IF(AND([@Reactive]<>"",[@Reactive]<>"Lag",[@Reactive]<>"Lead"),"ERR: Reactive must be Lag or Lead",'
           'IF(AND([@ExclGroup]<>"",COUNTIFS(tblLoad[ExclGroup],[@ExclGroup],tblLoad[TX],"<>"&[@TX])+COUNTIFS(tblLoad[ExclGroup],[@ExclGroup],tblLoad[Feeder],"<>"&[@Feeder])>0),"ERR: ExclGroup members must share the same TX and Feeder",'
           'IF(AND([@TX]="",[@Phase_e]<>"DC"),"WARN: no transformer assigned (site totals only)",'
           'IF([@Cat_e]="","WARN: category missing",'
           'IF(ISNA(MATCH([@Cat_e],tblCat[Category],0)),"WARN: category not in tblCat",'
           'IF(AND([@Phase_e]="1PH",[@Conn1Ph]=""),"WARN: 1PH connection blank - L-N assumed",'
           'IF([@Phase_e]="DC","INFO: DC load - excluded from AC totals (supply via charger row)",'
           'IF(OR(NOT(ISNUMBER([@DF])),NOT(ISNUMBER([@DutyF])),NOT(ISNUMBER([@Hours]))),"INFO: design-basis default DF/duty/hours applied",'
           'IF([@Reference]="","INFO: no datasheet reference","OK")))))))))))))))))))))))))))')
    cols[col_index(cols, "Check")]["f"] = chk
    return cols


def build_load(wb, ws, cfg):
    title(ws, "Sheet 3 - Dynamic Auxiliary Load Schedule (tblLoad)",
          "Enter one row per equipment type per supply point. Type below the last row (or Tab in the last cell) to add rows - "
          "formulas extend automatically. Library values fill any blank yellow cell where a LibID is selected.")
    ws["A3"] = ("Equations: 3PH S=√3·V_LL·I/1000; 1PH/DC S=V·I/1000; P=S·PF; kW_out: P=Pout/η; Connected = QtyInst×P; "
                "Demand = QtyDuty×P×DF; Scenario = Demand×profile factor; Energy = Demand×DutyF×Hours. See sheet 10.")
    ws["A3"].font = FN_NOTE
    cols = load_cols()
    rows = cfg.get("loads", [])
    f, l = make_table(wb, ws, "tblLoad", cols, rows, top=6, left=1, nblank=cfg.get("load_blank", 40))
    ci = col_index(cols, "Check") + 1
    status_cf(ws, f"{CL(ci)}7:{CL(ci)}{l + 300}")
    ws.freeze_panes = "D7"
    ws.cell(row=4, column=1, value="INPUTS →").font = FN_B
    ws.cell(row=4, column=ci, value="CALCULATIONS (do not overwrite) →").font = FN_B
    return cols


# ======================================================================= TRANSFORMERS
def tx_cols():
    cols = [
        dict(h="TX_ID", kind="in", w=9, unit="Transformer / supply group ID"),
        dict(h="Description", kind="in", w=26, unit="Description"),
        dict(h="Section", kind="in", w=14, unit="LV board / section supplied"),
        dict(h="Role", kind="in", w=14, unit="Arrangement role", dv="=lstRole"),
        dict(h="RedGroup", kind="in", w=8, unit="Redundancy group (blank = none)"),
        dict(h="Scheme", kind="in", w=7, unit="N / N+1 / 2N (blank = N)", dv="=lstScheme"),
        dict(h="HV_kV", kind="in", w=7, fmt="0.0##", unit="HV kV (blank = DB)", dvnum=(0.1, 1000)),
        dict(h="LV_VLL", kind="in", w=7, fmt="0", unit="LV V L-L (blank = DB)", dvnum=(50, 50000)),
        dict(h="LV_VLN", kind="in", w=7, fmt="0", unit="LV V L-N (blank = VLL/√3)", dvnum=(20, 50000)),
        dict(h="Type", kind="in", w=12, unit="Liquid-filled / Dry-type", dv="=lstTXType"),
        dict(h="Cooling", kind="in", w=10, unit="Depends on type", dv=None),
        dict(h="Z_pct", kind="in", w=6, fmt="0.00", unit="Impedance uk (%) at rated kVA", dvnum=(0.5, 30)),
        dict(h="XR", kind="in", w=6, fmt="0.0", unit="X/R ratio", dvnum=(0.1, 100)),
        dict(h="Kharm_in", kind="in", w=7, fmt="0.000", unit="Harmonic derating factor (≤1) per IEEE C57.110 / mfr", dvnum=(0.1, 1)),
        dict(h="InrushMult", kind="in", w=7, fmt="0.0", unit="Energisation inrush ÷ I_HV rated (mfr)", dvnum=(0, 50)),
        dict(h="DesignScen", kind="in", w=10, unit="Blank = DB default", dv="=lstDesignScen"),
        dict(h="Selected_kVA", kind="in", w=9, fmt="#,##0", unit="Engineer-selected rating (kVA)", dvnum=(1, 1e6)),
        dict(h="Notes", kind="in", w=24, unit="Notes"),
        dict(h="Check", kind="calc", w=40, unit="Transformer check (first issue)"),
        dict(h="HVkV_e", kind="calc", w=6, fmt="0.0##", f='IF(ISNUMBER([@HV_kV]),[@HV_kV],N(DB_MV))', unit="HV kV"),
        dict(h="VLL_e", kind="calc", w=6, fmt="0", f='IF(ISNUMBER([@LV_VLL]),[@LV_VLL],N(DB_VLL))', unit="LV V L-L"),
        dict(h="VLN_e", kind="calc", w=6, fmt="0", f='IF(ISNUMBER([@LV_VLN]),[@LV_VLN],[@VLL_e]/SQRT3)', unit="LV V L-N"),
        dict(h="Scen_e", kind="calc", w=9, f='IF([@DesignScen]<>"",[@DesignScen],DB_DesignScen)', unit="Design scenario"),
    ]
    for s in SC:
        cols.append(dict(h=f"S_{s}", kind="calc", w=8, fmt="#,##0.0",
                         f=f'IF([@TX_ID]="",0,SQRT(SUMIFS(tblLoad[P_{s}],tblLoad[TX],[@TX_ID])^2+SUMPRODUCT((tblLoad[TX]=[@TX_ID])*tblLoad[P_{s}]*tblLoad[TanPhi])^2))',
                         unit=f"kVA {s}"))
    sl = ",".join(f"[@S_{s}]" for s in SC)
    gov = '"-"'
    for s in reversed(SC):
        gov = f'IF([@S_{s}]=[@Design_kVA],"{s}",{gov})'
    cols += [
        dict(h="Conn_kVA", kind="calc", w=8, fmt="#,##0.0", f='IF([@TX_ID]="",0,SUMIFS(tblLoad[Conn_kVA],tblLoad[TX],[@TX_ID]))', unit="Connected kVA (arith.)"),
        dict(h="Conn_kW", kind="calc", w=8, fmt="#,##0.0", f='IF([@TX_ID]="",0,SUMIFS(tblLoad[Conn_kW],tblLoad[TX],[@TX_ID]))', unit="Connected kW"),
        dict(h="P_run", kind="calc", w=8, fmt="#,##0.0", f='IF([@TX_ID]="",0,SUMIFS(tblLoad[Run_kW],tblLoad[TX],[@TX_ID]))', unit="Σ demand kW (no scenario)"),
        dict(h="Q_run", kind="calc", w=8, fmt="#,##0.0", f='IF([@TX_ID]="",0,SUMPRODUCT((tblLoad[TX]=[@TX_ID])*tblLoad[Run_kW]*tblLoad[TanPhi]))', unit="Σ demand kvar"),
        dict(h="PF_run", kind="calc", w=6, fmt="0.000", f='IF([@P_run]>0,[@P_run]/SQRT([@P_run]^2+[@Q_run]^2),"")', unit="Aggregate PF"),
        dict(h="Sin_run", kind="calc", w=6, fmt="0.000", f='IF([@P_run]>0,[@Q_run]/SQRT([@P_run]^2+[@Q_run]^2),0)', unit="Aggregate sin φ (signed)"),
        dict(h="Design_kVA", kind="calc", w=9, fmt="#,##0.0",
             f=f'IF([@Scen_e]="ENVELOPE",MAX({sl}),IFERROR(CHOOSE(MATCH([@Scen_e],lstScenCodes,0),{sl}),0))', unit="Maximum coincident demand (kVA)"),
        dict(h="Gov", kind="calc", w=8, f=f'IF([@Design_kVA]=0,"-",IF([@Scen_e]<>"ENVELOPE",[@Scen_e],{gov}))', unit="Governing scenario"),
        dict(h="Kamb", kind="calc", w=6, fmt="0.000",
             f=('IF([@Type]="Dry-type",IF(N(DB_Tamb)>DB_RefT_D,IF(ISNUMBER(DB_DerK_D),MAX(0,1-(DB_Tamb-DB_RefT_D)*DB_DerK_D/100),1),1),'
                'IF(N(DB_Tamb)>DB_RefT_L,IF(ISNUMBER(DB_DerK_L),MAX(0,1-(DB_Tamb-DB_RefT_L)*DB_DerK_L/100),1),1))'), unit="Ambient factor"),
        dict(h="Kalt", kind="calc", w=6, fmt="0.000",
             f=('IF([@Type]="Dry-type",IF(N(DB_Alt)>DB_RefAlt_D,IF(ISNUMBER(DB_DerAlt_D),MAX(0,1-(DB_Alt-DB_RefAlt_D)/100*DB_DerAlt_D/100),1),1),'
                'IF(N(DB_Alt)>DB_RefAlt_L,IF(ISNUMBER(DB_DerAlt_L),MAX(0,1-(DB_Alt-DB_RefAlt_L)/100*DB_DerAlt_L/100),1),1))'), unit="Altitude factor"),
        dict(h="Kharm", kind="calc", w=6, fmt="0.000", f='IF(ISNUMBER([@Kharm_in]),[@Kharm_in],1)', unit="Harmonic factor"),
        dict(h="Kderate", kind="calc", w=6, fmt="0.000", f='[@Kamb]*[@Kalt]*[@Kharm]', unit="Total derating"),
        dict(h="DesignPlus", kind="calc", w=9, fmt="#,##0.0", f='[@Design_kVA]*(1+N(DB_Margin))*(1+N(DB_Future))', unit="Demand × (1+margin)×(1+future)"),
        dict(h="Req_own", kind="calc", w=9, fmt="#,##0.0", f='IF([@Kderate]>0,[@DesignPlus]/[@Kderate],0)', unit="Required kVA (own load)"),
        dict(h="Scheme_e", kind="calc", w=6, f='IF([@Scheme]="","N",[@Scheme])', unit="Scheme"),
        dict(h="GrpN", kind="calc", w=5, f='IF([@RedGroup]="",1,COUNTIF(tblTX[RedGroup],[@RedGroup]))', unit="Units in group"),
        dict(h="GrpDesign", kind="calc", w=9, fmt="#,##0.0", f='IF([@RedGroup]="",[@Design_kVA],SUMIFS(tblTX[Design_kVA],tblTX[RedGroup],[@RedGroup]))', unit="Group demand kVA"),
        dict(h="GrpPlus", kind="calc", w=9, fmt="#,##0.0", f='IF([@RedGroup]="",[@DesignPlus],SUMIFS(tblTX[DesignPlus],tblTX[RedGroup],[@RedGroup]))', unit="Group demand + margins"),
        dict(h="UnitsAvail", kind="calc", w=6, fmt="0.##", f='IF([@Scheme_e]="N+1",[@GrpN]-1,IF([@Scheme_e]="2N",[@GrpN]/2,[@GrpN]))', unit="Units in service at contingency"),
        dict(h="Req_cont", kind="calc", w=9, fmt="#,##0.0", f='IF([@Scheme_e]="N",0,IF(AND([@UnitsAvail]>0,[@Kderate]>0),[@GrpPlus]/[@UnitsAvail]/[@Kderate],0))', unit="Required kVA at contingency"),
        dict(h="Req_min", kind="calc", w=9, fmt="#,##0.0", f='MAX([@Req_own],[@Req_cont])', unit="CALCULATED MINIMUM RATING (kVA)"),
        dict(h="Rec_std", kind="calc", w=10, fmt="#,##0.#",
             f='IF([@Req_min]<=0,"No load",IF(COUNT(lstStdActive)=0,"No std list",IF([@Req_min]>MAX(lstStdActive),"Exceeds std list",SMALL(lstStdActive,COUNTIF(lstStdActive,"<"&[@Req_min])+1))))',
             unit="RECOMMENDED STANDARD RATING (not approved)"),
        dict(h="Eval_kVA", kind="calc", w=9, fmt="#,##0.#", f='IF(ISNUMBER([@Selected_kVA]),[@Selected_kVA],IF(ISNUMBER([@Rec_std]),[@Rec_std],0))', unit="Rating used for evaluation"),
        dict(h="DerateOK", kind="calc", w=6,
             f=('IF([@TX_ID]="",1,IF(OR(AND(N(DB_Tamb)>IF([@Type]="Dry-type",DB_RefT_D,DB_RefT_L),NOT(ISNUMBER(IF([@Type]="Dry-type",DB_DerK_D,DB_DerK_L)))),'
                'AND(N(DB_Alt)>IF([@Type]="Dry-type",DB_RefAlt_D,DB_RefAlt_L),NOT(ISNUMBER(IF([@Type]="Dry-type",DB_DerAlt_D,DB_DerAlt_L)))),'
                'AND([@NL_share]>N(DB_NLThr),NOT(ISNUMBER([@Kharm_in])))),0,1))'),
             unit="1 = derating basis complete (ambient, altitude, harmonics); 0 = sizing cannot be confirmed"),
        dict(h="RatingStatus", kind="calc", w=30,
             f=('IF([@TX_ID]="","",IF(ISNUMBER([@Selected_kVA]),IF([@Selected_kVA]<[@Req_min]-0.000001,"FAIL: SELECTED below calculated minimum",'
                'IF([@DerateOK]=0,"REVIEW: selected - derating data incomplete, sizing not confirmed","SELECTED - adequate")),'
                'IF(ISNUMBER([@Rec_std]),IF([@DerateOK]=0,"PROVISIONAL - derating data incomplete","PROVISIONAL - recommended std, not approved"),"REVIEW: no rating available")))'),
             unit="Selection status"),
        dict(h="Cap_derated", kind="calc", w=9, fmt="#,##0.0", f='[@Eval_kVA]*[@Kderate]', unit="Derated capacity kVA"),
        dict(h="Load_pct", kind="calc", w=7, fmt="0.0", f='IF([@Eval_kVA]>0,[@Design_kVA]/[@Eval_kVA]*100,0)', unit="Loading % of rating"),
        dict(h="LoadDer_pct", kind="calc", w=7, fmt="0.0", f='IF([@Cap_derated]>0,[@Design_kVA]/[@Cap_derated]*100,0)', unit="Loading % of derated capacity"),
        dict(h="Spare_kVA", kind="calc", w=9, fmt="#,##0.0", f='IF([@TX_ID]="","",[@Cap_derated]-[@Design_kVA])', unit="Spare kVA (derated − demand)"),
        dict(h="Spare_pct", kind="calc", w=7, fmt="0.0", f='IF(N([@Cap_derated])>0,[@Spare_kVA]/[@Cap_derated]*100,0)', unit="Spare %"),
        dict(h="ContLoad_pct", kind="calc", w=8, fmt="0.0", f='IF([@Scheme_e]="N","",IF(AND([@UnitsAvail]>0,[@Cap_derated]>0),[@GrpDesign]/[@UnitsAvail]/[@Cap_derated]*100,""))', unit="Loading % at contingency"),
    ]
    for s in SC:
        cols.append(dict(h=f"U_{s}", kind="calc", w=7, fmt="0.0", f=f'IF([@Eval_kVA]>0,[@S_{s}]/[@Eval_kVA]*100,0)', unit=f"Utilisation % {s}"))
    cols += [
        dict(h="I_HV", kind="calc", w=7, fmt="#,##0.0", f='IF(AND([@Eval_kVA]>0,[@HVkV_e]>0),[@Eval_kVA]/(SQRT3*[@HVkV_e]),0)', unit="HV full-load A"),
        dict(h="I_LV", kind="calc", w=8, fmt="#,##0.0", f='IF(AND([@Eval_kVA]>0,[@VLL_e]>0),[@Eval_kVA]*1000/(SQRT3*[@VLL_e]),0)', unit="LV full-load A"),
        dict(h="I_LV_design", kind="calc", w=8, fmt="#,##0.0", f='IF([@VLL_e]>0,[@Design_kVA]*1000/(SQRT3*[@VLL_e]),0)', unit="LV design current A"),
        dict(h="ur", kind="calc", w=6, fmt="0.000", f='IF(AND(ISNUMBER([@Z_pct]),ISNUMBER([@XR])),[@Z_pct]/SQRT(1+[@XR]^2),"")', unit="Resistive uR %"),
        dict(h="ux", kind="calc", w=6, fmt="0.000", f='IF(ISNUMBER([@ur]),[@ur]*[@XR],"")', unit="Reactive uX %"),
        dict(h="Reg_pct", kind="calc", w=7, fmt="0.00",
             f=('IF(AND(ISNUMBER([@ur]),[@Eval_kVA]>0,ISNUMBER([@PF_run])),([@Design_kVA]/[@Eval_kVA])*([@ur]*[@PF_run]+[@ux]*[@Sin_run])'
                '+(([@Design_kVA]/[@Eval_kVA])*([@ux]*[@PF_run]-[@ur]*[@Sin_run]))^2/200,"")'), unit="Voltage regulation % at design"),
        dict(h="Ssc_tx", kind="calc", w=9, fmt="#,##0", f='IF(AND(ISNUMBER([@Z_pct]),[@Eval_kVA]>0),[@Eval_kVA]/([@Z_pct]/100),"")', unit="TX short-circuit kVA"),
        dict(h="Ssc_bus", kind="calc", w=9, fmt="#,##0", f='IF(ISNUMBER([@Ssc_tx]),IF(N(DB_SrcMVA)>0,1/(1/[@Ssc_tx]+1/(DB_SrcMVA*1000)),[@Ssc_tx]),"")', unit="LV bus short-circuit kVA"),
        dict(h="Isc_kA", kind="calc", w=7, fmt="0.00", f='IF(AND(ISNUMBER([@Ssc_bus]),ISNUMBER(DB_Cmax),[@VLL_e]>0),DB_Cmax*[@Ssc_bus]/(SQRT3*[@VLL_e]),"")', unit="Prospective LV 3-ph Isc kA (prelim.)"),
        dict(h="MaxStart_kVA", kind="calc", w=8, fmt="#,##0.0", f='IF([@TX_ID]="",0,SUMPRODUCT(MAX((tblLoad[TX]=[@TX_ID])*tblLoad[Start_kVA])))', unit="Largest single starting kVA"),
        dict(h="Dip_pct", kind="calc", w=6, fmt="0.00", f='IF(AND(ISNUMBER([@Ssc_bus]),[@MaxStart_kVA]>0),[@MaxStart_kVA]/([@MaxStart_kVA]+[@Ssc_bus])*100,"")', unit="Start voltage dip % (approx.)"),
        dict(h="NL_share", kind="calc", w=6, fmt="0.0%", f='IF([@TX_ID]="",0,IF(SUMIFS(tblLoad[Run_kVA],tblLoad[TX],[@TX_ID])>0,SUMIFS(tblLoad[Run_kVA],tblLoad[TX],[@TX_ID],tblLoad[NL_e],"Y")/SUMIFS(tblLoad[Run_kVA],tblLoad[TX],[@TX_ID]),0))', unit="Non-linear share of demand kVA"),
        dict(h="Inrush_A", kind="calc", w=8, fmt="#,##0", f='IF(ISNUMBER([@InrushMult]),[@InrushMult]*[@I_HV],"")', unit="Energisation inrush A (HV)"),
        dict(h="L1_kVA", kind="calc", w=7, fmt="#,##0.00", f='IF([@TX_ID]="",0,SUMIFS(tblLoad[Run_kVA],tblLoad[TX],[@TX_ID],tblLoad[Phase_e],"3PH")/3+SUMIFS(tblLoad[Run_kVA],tblLoad[TX],[@TX_ID],tblLoad[Phase_e],"1PH",tblLoad[PhaseAssign],"L1")+(SUMIFS(tblLoad[Run_kVA],tblLoad[TX],[@TX_ID],tblLoad[Phase_e],"1PH",tblLoad[PhaseAssign],"L1-L2")+SUMIFS(tblLoad[Run_kVA],tblLoad[TX],[@TX_ID],tblLoad[Phase_e],"1PH",tblLoad[PhaseAssign],"L3-L1"))/2)', unit="Phase L1 demand kVA"),
        dict(h="L2_kVA", kind="calc", w=7, fmt="#,##0.00", f='IF([@TX_ID]="",0,SUMIFS(tblLoad[Run_kVA],tblLoad[TX],[@TX_ID],tblLoad[Phase_e],"3PH")/3+SUMIFS(tblLoad[Run_kVA],tblLoad[TX],[@TX_ID],tblLoad[Phase_e],"1PH",tblLoad[PhaseAssign],"L2")+(SUMIFS(tblLoad[Run_kVA],tblLoad[TX],[@TX_ID],tblLoad[Phase_e],"1PH",tblLoad[PhaseAssign],"L1-L2")+SUMIFS(tblLoad[Run_kVA],tblLoad[TX],[@TX_ID],tblLoad[Phase_e],"1PH",tblLoad[PhaseAssign],"L2-L3"))/2)', unit="Phase L2 demand kVA"),
        dict(h="L3_kVA", kind="calc", w=7, fmt="#,##0.00", f='IF([@TX_ID]="",0,SUMIFS(tblLoad[Run_kVA],tblLoad[TX],[@TX_ID],tblLoad[Phase_e],"3PH")/3+SUMIFS(tblLoad[Run_kVA],tblLoad[TX],[@TX_ID],tblLoad[Phase_e],"1PH",tblLoad[PhaseAssign],"L3")+(SUMIFS(tblLoad[Run_kVA],tblLoad[TX],[@TX_ID],tblLoad[Phase_e],"1PH",tblLoad[PhaseAssign],"L2-L3")+SUMIFS(tblLoad[Run_kVA],tblLoad[TX],[@TX_ID],tblLoad[Phase_e],"1PH",tblLoad[PhaseAssign],"L3-L1"))/2)', unit="Phase L3 demand kVA"),
        dict(h="Unalloc_kVA", kind="calc", w=7, fmt="#,##0.00", f='IF([@TX_ID]="",0,SUMIFS(tblLoad[Run_kVA],tblLoad[TX],[@TX_ID],tblLoad[Phase_e],"1PH",tblLoad[PhaseAssign],""))', unit="1PH kVA without phase allocation"),
        dict(h="Imb_pct", kind="calc", w=6, fmt="0.00", f='IF([@L1_kVA]+[@L2_kVA]+[@L3_kVA]>0,(MAX([@L1_kVA],[@L2_kVA],[@L3_kVA])-([@L1_kVA]+[@L2_kVA]+[@L3_kVA])/3)/(([@L1_kVA]+[@L2_kVA]+[@L3_kVA])/3)*100,0)', unit="Imbalance % (max dev ÷ avg)"),
        dict(h="BalCheck", kind="calc", w=14, f='IF([@TX_ID]="","",IF([@Unalloc_kVA]>0,"REVIEW: 1PH loads without phase allocation",IF(NOT(ISNUMBER(DB_ImbLim)),"REVIEW: limit not set",IF([@Imb_pct]<=DB_ImbLim*100,"PASS","FAIL"))))', unit="Phase balance check"),
    ]
    chk = ('IF([@TX_ID]="","",IF(COUNTIF(tblTX[TX_ID],[@TX_ID])>1,"ERR: duplicate TX_ID",'
           'IF([@HVkV_e]<=0,"ERR: HV voltage missing",IF([@VLL_e]<=0,"ERR: LV voltage missing",'
           'IF(AND([@Type]<>"Liquid-filled",[@Type]<>"Dry-type"),"ERR: transformer type missing",'
           'IF(AND([@Scheme_e]="N+1",[@GrpN]<2),"ERR: N+1 needs >=2 units in RedGroup",'
           'IF(AND([@Scheme_e]="2N",OR([@GrpN]<2,MOD([@GrpN],2)<>0)),"ERR: 2N needs an even number of units in RedGroup",'
           'IF(AND([@Scheme_e]<>"N",[@RedGroup]=""),"ERR: redundancy scheme needs a RedGroup",'
           'IF(AND([@RedGroup]<>"",COUNTIFS(tblTX[RedGroup],[@RedGroup],tblTX[Scheme_e],[@Scheme_e])<>[@GrpN]),"ERR: mixed schemes in one RedGroup",'
           'IF(LEFT([@RatingStatus],4)="FAIL",[@RatingStatus],'
           'IF([@LoadDer_pct]>DB_MaxLoad*100+0.000001,"FAIL: loading exceeds permissible",'
           'IF(AND([@Scheme_e]<>"N",ISNUMBER([@ContLoad_pct]),N([@ContLoad_pct])>DB_MaxLoadCont*100+0.000001),"FAIL: contingency loading exceeds permissible",'
           'IF(AND(N(DB_Tamb)>IF([@Type]="Dry-type",DB_RefT_D,DB_RefT_L),NOT(ISNUMBER(IF([@Type]="Dry-type",DB_DerK_D,DB_DerK_L)))),"REVIEW: ambient above rating basis - derating coefficient not entered",'
           'IF(AND(N(DB_Alt)>IF([@Type]="Dry-type",DB_RefAlt_D,DB_RefAlt_L),NOT(ISNUMBER(IF([@Type]="Dry-type",DB_DerAlt_D,DB_DerAlt_L)))),"REVIEW: altitude above rating basis - derating not entered",'
           'IF(AND(ISNUMBER(DB_NLThr),[@NL_share]>DB_NLThr,NOT(ISNUMBER([@Kharm_in]))),"REVIEW: non-linear share above threshold - harmonic derating not entered",'
           'IF(OR(NOT(ISNUMBER([@Z_pct])),NOT(ISNUMBER([@XR]))),"REVIEW: impedance / X/R not entered",'
           'IF([@Cooling]="","WARN: cooling not stated",'
           'IF(NOT(ISNUMBER([@Selected_kVA])),"REVIEW: rating not selected (provisional)","OK"))))))))))))))))))')
    cols[col_index(cols, "Check")]["f"] = chk
    return cols


def build_tx(wb, ws, cfg):
    title(ws, "Sheet 6 - Auxiliary Transformer Sizing (tblTX)",
          "One row per auxiliary transformer (centralized = one row / one N+1 group; distributed = one row per section). "
          "Required kVA = Demand×(1+margin)×(1+future) ÷ (Kamb·Kalt·Kharm); N+1 / 2N: group demand ÷ units in service.")
    ws["A3"] = ("Calculated minimum rating (Req_min) and recommended standard rating (Rec_std) are separate. Rec_std is NOT an "
                "approved selection - enter the engineer-selected rating in Selected_kVA after manufacturer / client review.")
    ws["A3"].font = Font(bold=True, color="C00000")
    cols = tx_cols()
    f, l = make_table(wb, ws, "tblTX", cols, cfg.get("tx", []), top=6, left=1, nblank=cfg.get("tx_blank", 2))
    add_name(wb, "lstTX", "tblTX[TX_ID]")
    ci = col_index(cols, "Check") + 1
    status_cf(ws, f"{CL(ci)}7:{CL(ci)}{l + 100}")
    rs = col_index(cols, "RatingStatus") + 1
    status_cf(ws, f"{CL(rs)}7:{CL(rs)}{l + 100}")
    ws.freeze_panes = "B7"
    # Fix dependent cooling DV: replace INDIRECT R1C1 hack with explicit A1 reference
    tcol = CL(col_index(cols, "Type") + 1)
    ccol = CL(col_index(cols, "Cooling") + 1)
    for dv in list(ws.data_validations.dataValidation):
        if "INDIRECT" in (dv.formula1 or ""):
            ws.data_validations.dataValidation.remove(dv)
    dv = DataValidation(type="list", formula1=f'=IF(${tcol}7="Dry-type",lstCoolDry,lstCoolLiq)', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{ccol}7:{ccol}{l + 100}")
    # centralized vs distributed comparison block
    r0 = l + 4
    ws.cell(row=r0, column=1, value="Centralized vs distributed comparison (site level)").font = FN_B
    items = [
        ("Site coincident design demand (kVA) - all AC loads, peak scenario", "=MAX(SCN_S)"),
        ("Σ of per-transformer design demands (kVA) - non-coincident", "=SUM(tblTX[Design_kVA])"),
        ("Diversity between sections (Σ per-TX ÷ site coincident)", '=IF(C{a}>0,C{b}/C{a},"")'),
        ("Centralized comparison: transformer type", None),
        ("Centralized comparison: ambient×altitude factor (harmonic factor excluded)",
         '=IF(C{d}="Dry-type",IF(N(DB_Tamb)>DB_RefT_D,IF(ISNUMBER(DB_DerK_D),MAX(0,1-(DB_Tamb-DB_RefT_D)*DB_DerK_D/100),1),1)*IF(N(DB_Alt)>DB_RefAlt_D,IF(ISNUMBER(DB_DerAlt_D),MAX(0,1-(DB_Alt-DB_RefAlt_D)/100*DB_DerAlt_D/100),1),1),'
         'IF(N(DB_Tamb)>DB_RefT_L,IF(ISNUMBER(DB_DerK_L),MAX(0,1-(DB_Tamb-DB_RefT_L)*DB_DerK_L/100),1),1)*IF(N(DB_Alt)>DB_RefAlt_L,IF(ISNUMBER(DB_DerAlt_L),MAX(0,1-(DB_Alt-DB_RefAlt_L)/100*DB_DerAlt_L/100),1),1))'),
        ("Centralized single transformer: calculated minimum (kVA)", '=IF(C{e}>0,C{a}*(1+N(DB_Margin))*(1+N(DB_Future))/C{e},0)'),
        ("Centralized single transformer: recommended standard (kVA)", '=IF(C{f}<=0,"No load",IF(C{f}>MAX(lstStdActive),"Exceeds std list",SMALL(lstStdActive,COUNTIF(lstStdActive,"<"&C{f})+1)))'),
        ("Centralized 2 × 100 % (N+1): each unit minimum (kVA)", "=C{f}"),
        ("Distributed: Σ evaluated ratings installed (kVA)", "=SUM(tblTX[Eval_kVA])"),
        ("Distributed: number of transformers", '=COUNTIF(tblTX[TX_ID],"?*")'),
    ]
    rowmap = {}
    keys = "abcdefghij"
    for i, (lab, fx) in enumerate(items):
        rowmap[keys[i]] = r0 + 1 + i
    for i, (lab, fx) in enumerate(items):
        r = r0 + 1 + i
        ws.cell(row=r, column=1, value=lab)
        c = ws.cell(row=r, column=3)
        if fx is None:
            c.value = cfg.get("central_type", "Liquid-filled")
            c.fill = F_IN
            c.font = FN_IN
            d = DataValidation(type="list", formula1="=lstTXType")
            ws.add_data_validation(d)
            d.add(f"C{r}")
        else:
            c.value = fx.format(**rowmap)
            c.fill = F_CALC
            c.number_format = "#,##0.00"
    add_name(wb, "CMP_SiteDesign", f"'Transformer Sizing'!$C${rowmap['a']}")
    add_name(wb, "CMP_SumTX", f"'Transformer Sizing'!$C${rowmap['b']}")
    add_name(wb, "CMP_CentMin", f"'Transformer Sizing'!$C${rowmap['f']}")
    add_name(wb, "CMP_CentRec", f"'Transformer Sizing'!$C${rowmap['g']}")
    add_name(wb, "CMP_DistInst", f"'Transformer Sizing'!$C${rowmap['i']}")
    ws.cell(row=r0 + 12, column=1, value=(
        "Notes: (1) Contingency sizing assumes the group load redistributes equally to the units in service (bus-tie / "
        "auto-transfer scheme must achieve this). (2) Loading above 100 % at contingency requires a loading-guide study "
        "(IEC 60076-7 / IEEE C57.91 for liquid; IEC 60076-11/-12 / IEEE C57.96 for dry-type). (3) Voltage regulation uses the "
        "approximate formula ε = n(uR·cosφ + uX·sinφ) + n²(uX·cosφ − uR·sinφ)²/200. (4) Isc is a preliminary symmetrical "
        "value: c·S_sc/(√3·U); a full IEC 60909 / IEEE study is required for equipment ratings.")).font = FN_NOTE
    return cols


# ======================================================================= FEEDERS
def fdr_cols():
    cols = [
        dict(h="FeederID", kind="in", w=10, unit="Feeder / circuit ID"),
        dict(h="Board", kind="in", w=12, unit="LV board / section"),
        dict(h="TX_ID", kind="in", w=8, unit="Supply transformer", dv="=lstTX"),
        dict(h="Description", kind="in", w=26, unit="Description"),
        dict(h="Phase", kind="in", w=6, unit="3PH / 1PH", dv="=lstPhase"),
        dict(h="Conn", kind="in", w=6, unit="1PH: L-N / L-L", dv="=lstConn"),
        dict(h="Basis", kind="in", w=10, unit="Design scenario (blank = ENVELOPE)", dv="=lstDesignScen"),
        dict(h="Length_m", kind="in", w=7, fmt="0", unit="Route length m", dvnum=(0, 100000)),
        dict(h="Cable", kind="in", w=16, unit="Cable description (input only)"),
        dict(h="Iz_A", kind="in", w=7, fmt="0", unit="Derated cable capacity Iz (A) from standard tables × correction factors", dvnum=(0, 100000)),
        dict(h="R_ohm_km", kind="in", w=7, fmt="0.0000", unit="Conductor R at operating temp Ω/km", dvnum=(0, 100)),
        dict(h="X_ohm_km", kind="in", w=7, fmt="0.0000", unit="Reactance Ω/km", dvnum=(0, 100)),
        dict(h="In_A", kind="in", w=7, fmt="0", unit="Protective device rating In (A)", dvnum=(0, 100000)),
        dict(h="I2_ratio", kind="in", w=6, fmt="0.00", unit="I2 ÷ In (device standard)", dvnum=(1, 3)),
        dict(h="Icu_kA", kind="in", w=6, fmt="0.0", unit="Device breaking capacity kA", dvnum=(0, 500)),
        dict(h="Notes", kind="in", w=20, unit="Notes"),
        dict(h="Overall", kind="calc", w=26, unit="Overall feeder check"),
        dict(h="VLL", kind="calc", w=6, fmt="0", f='IFERROR(INDEX(tblTX[VLL_e],MATCH([@TX_ID],tblTX[TX_ID],0)),N(DB_VLL))', unit="V L-L"),
        dict(h="VLN", kind="calc", w=6, fmt="0", f='IFERROR(INDEX(tblTX[VLN_e],MATCH([@TX_ID],tblTX[TX_ID],0)),N(DB_VLN))', unit="V L-N"),
        dict(h="V_fdr", kind="calc", w=6, fmt="0", f='IF([@Phase]="1PH",IF([@Conn]="L-L",[@VLL],[@VLN]),[@VLL])', unit="Feeder V"),
        dict(h="Basis_e", kind="calc", w=9, f='IF([@Basis]="","ENVELOPE",[@Basis])', unit="Basis"),
    ]
    for s in SC:
        cols.append(dict(h=f"S_{s}", kind="calc", w=7, fmt="#,##0.0",
                         f=f'IF([@FeederID]="",0,SQRT(SUMIFS(tblLoad[P_{s}],tblLoad[Feeder],[@FeederID])^2+SUMPRODUCT((tblLoad[Feeder]=[@FeederID])*tblLoad[P_{s}]*tblLoad[TanPhi])^2))',
                         unit=f"kVA {s}"))
    sl = ",".join(f"[@S_{s}]" for s in SC)
    cols += [
        dict(h="Conn_kVA", kind="calc", w=8, fmt="#,##0.0", f='IF([@FeederID]="",0,SUMIFS(tblLoad[Conn_kVA],tblLoad[Feeder],[@FeederID]))', unit="Connected kVA"),
        dict(h="Loads", kind="calc", w=5, f='IF([@FeederID]="",0,COUNTIFS(tblLoad[Feeder],[@FeederID],tblLoad[Valid],1))', unit="Load rows"),
        dict(h="P_run", kind="calc", w=8, fmt="#,##0.0", f='IF([@FeederID]="",0,SUMIFS(tblLoad[Run_kW],tblLoad[Feeder],[@FeederID]))', unit="Σ kW"),
        dict(h="Q_run", kind="calc", w=8, fmt="#,##0.0", f='IF([@FeederID]="",0,SUMPRODUCT((tblLoad[Feeder]=[@FeederID])*tblLoad[Run_kW]*tblLoad[TanPhi]))', unit="Σ kvar"),
        dict(h="PF_run", kind="calc", w=6, fmt="0.000", f='IF([@P_run]>0,[@P_run]/SQRT([@P_run]^2+[@Q_run]^2),1)', unit="Aggregate PF"),
        dict(h="Sin_run", kind="calc", w=6, fmt="0.000", f='IF([@P_run]>0,[@Q_run]/SQRT([@P_run]^2+[@Q_run]^2),0)', unit="Aggregate sin φ (signed)"),
        dict(h="S_design", kind="calc", w=8, fmt="#,##0.0", f=f'IF([@Basis_e]="ENVELOPE",MAX({sl}),IFERROR(CHOOSE(MATCH([@Basis_e],lstScenCodes,0),{sl}),0))', unit="Design kVA"),
        dict(h="Ib_A", kind="calc", w=7, fmt="#,##0.0", f='IF([@V_fdr]>0,IF([@Phase]="1PH",[@S_design]*1000/[@V_fdr],[@S_design]*1000/(SQRT3*[@V_fdr])),0)', unit="Design current Ib (A)"),
        dict(h="In_min", kind="calc", w=7, fmt="#,##0.0", f='[@Ib_A]*DB_ContMult', unit="Minimum In (A)"),
        dict(h="Chk_In", kind="calc", w=10, f='IF(NOT(ISNUMBER([@In_A])),"REVIEW: In not entered",IF([@In_A]>=[@In_min],"PASS","FAIL"))', unit="In ≥ Ib×mult"),
        dict(h="Chk_Iz", kind="calc", w=10, f='IF(NOT(ISNUMBER([@Iz_A])),"REVIEW: Iz not entered",IF(NOT(ISNUMBER([@In_A])),"REVIEW: In not entered",IF([@Iz_A]>=[@In_A],"PASS","FAIL")))', unit="Iz ≥ In"),
        dict(h="Chk_I2", kind="calc", w=10, f='IF(OR(NOT(ISNUMBER([@I2_ratio])),NOT(ISNUMBER([@Iz_A])),NOT(ISNUMBER([@In_A]))),"REVIEW: data missing",IF([@I2_ratio]*[@In_A]<=DB_K145*[@Iz_A],"PASS","FAIL"))', unit="I2 ≤ K·Iz"),
        dict(h="VD_pct", kind="calc", w=6, fmt="0.00", f='IF(AND(ISNUMBER([@Length_m]),ISNUMBER([@R_ohm_km]),ISNUMBER([@X_ohm_km]),[@V_fdr]>0),IF([@Phase]="1PH",2,SQRT3)*[@Ib_A]*[@Length_m]/1000*([@R_ohm_km]*[@PF_run]+[@X_ohm_km]*[@Sin_run])/[@V_fdr]*100,"")', unit="Voltage drop %"),
        dict(h="Chk_VD", kind="calc", w=10, f='IF(NOT(ISNUMBER([@VD_pct])),"REVIEW: cable data missing",IF(NOT(ISNUMBER(DB_VDLim)),"REVIEW: VD limit not set",IF([@VD_pct]<=DB_VDLim*100,"PASS","FAIL")))', unit="VD ≤ limit"),
        dict(h="MaxStart_kVA", kind="calc", w=8, fmt="#,##0.0", f='IF([@FeederID]="",0,SUMPRODUCT(MAX((tblLoad[Feeder]=[@FeederID])*tblLoad[Start_kVA])))', unit="Largest start kVA"),
        dict(h="IStart_A", kind="calc", w=8, fmt="#,##0.0", f='IF([@V_fdr]>0,IF([@Phase]="1PH",[@MaxStart_kVA]*1000/[@V_fdr],[@MaxStart_kVA]*1000/(SQRT3*[@V_fdr])),0)', unit="Largest start current A (for device curve check)"),
        dict(h="Isc_bus_kA", kind="calc", w=7, fmt="0.00", f='IFERROR(INDEX(tblTX[Isc_kA],MATCH([@TX_ID],tblTX[TX_ID],0)),"")', unit="Isc at TX LV terminals (kA)"),
        dict(h="Chk_Icu", kind="calc", w=10, f='IF(OR(NOT(ISNUMBER([@Icu_kA])),NOT(ISNUMBER([@Isc_bus_kA]))),"REVIEW: data missing",IF([@Icu_kA]>=[@Isc_bus_kA],"PASS","FAIL"))', unit="Icu ≥ Isc"),
        dict(h="Load_pct", kind="calc", w=6, fmt="0.0", f='IF(N([@In_A])>0,[@Ib_A]/[@In_A]*100,"")', unit="Ib ÷ In %"),
    ]
    ov = ('IF([@FeederID]="","",IF(COUNTIF(tblFdr[FeederID],[@FeederID])>1,"ERR: duplicate FeederID",'
          'IF(ISNA(MATCH([@TX_ID],tblTX[TX_ID],0)),"ERR: TX_ID not defined",'
          'IF(AND([@Phase]<>"3PH",[@Phase]<>"1PH"),"ERR: phase must be 3PH or 1PH",'
          'IF(OR([@Chk_In]="FAIL",[@Chk_Iz]="FAIL",[@Chk_I2]="FAIL",[@Chk_VD]="FAIL",[@Chk_Icu]="FAIL"),"FAIL: see check columns",'
          'IF(OR(LEFT([@Chk_In],6)="REVIEW",LEFT([@Chk_Iz],6)="REVIEW",LEFT([@Chk_I2],6)="REVIEW",LEFT([@Chk_VD],6)="REVIEW",LEFT([@Chk_Icu],6)="REVIEW"),"REVIEW: incomplete data",'
          '"PASS (preliminary)"))))))')
    cols[col_index(cols, "Overall")]["f"] = ov
    return cols


def build_fdr(wb, ws, cfg):
    title(ws, "Sheet 7 - LV Auxiliary Distribution (tblFdr)",
          "PRELIMINARY CHECKS ONLY. Cable capacity Iz, R, X and device data are user inputs from the applicable standard "
          "(e.g. IEC 60364-5-52, NEC Art. 310) and manufacturer data. This sheet does not select cables or breakers.")
    ws["A3"] = ("Ib = S·1000/(√3·V) [3PH] or S·1000/V [1PH];  ΔV% = k·Ib·L·(R·cosφ + X·sinφ)/V·100, k = √3 (3PH) or 2 (1PH);  "
                "Coordination: Ib·mult ≤ In ≤ Iz and I2 ≤ K·Iz (K per Design Basis).")
    ws["A3"].font = FN_NOTE
    cols = fdr_cols()
    f, l = make_table(wb, ws, "tblFdr", cols, cfg.get("fdr", []), top=6, left=1, nblank=cfg.get("fdr_blank", 4))
    add_name(wb, "lstFdr", "tblFdr[FeederID]")
    for h in ("Overall", "Chk_In", "Chk_Iz", "Chk_I2", "Chk_VD", "Chk_Icu"):
        ci = col_index(cols, h) + 1
        status_cf(ws, f"{CL(ci)}7:{CL(ci)}{l + 100}")
    ws.freeze_panes = "B7"
    ws.cell(row=l + 3, column=1, value="Phase-load balance is calculated per transformer in tblTX (sheet 6, columns L1_kVA ... BalCheck) and scales with any number of transformers.").font = FN_NOTE
    return cols


# ======================================================================= SUMMARY
def build_summary(wb, ws, cfg):
    title(ws, "Sheet 5 - Auxiliary Load Summary", "All values calculated. Breakdown scenario is selectable (yellow cell).")
    ws.column_dimensions["A"].width = 52
    for c in "BCDEFGHI":
        ws.column_dimensions[c].width = 15
    kp = [
        ("Connected load - AC (kW)", "=SUM(tblLoad[Conn_kW])", "SUM_ConnkW", "#,##0.0"),
        ("Connected load - AC, arithmetic (kVA)", "=SUM(tblLoad[Conn_kVA])", "SUM_ConnkVA", "#,##0.0"),
        ("Connected load - AC, vector (kVA)", "=SQRT(SUM(tblLoad[Conn_kW])^2+SUMPRODUCT(tblLoad[Conn_kW],tblLoad[TanPhi])^2)", "SUM_ConnkVAv", "#,##0.0"),
        ("Demand before scenario factors (Σ duty×DF) (kW)", "=SUM(tblLoad[Run_kW])", "SUM_RunkW", "#,##0.0"),
        ("Maximum coincident demand - peak scenario (kVA)", "=MAX(SCN_S)", "SUM_PeakkVA", "#,##0.0"),
        ("Maximum coincident demand - peak scenario (kW)", "=INDEX(SCN_P,MATCH(MAX(SCN_S),SCN_S,0))", "SUM_PeakkW", "#,##0.0"),
        ("Power factor at peak scenario", '=IF(SUM_PeakkVA>0,SUM_PeakkW/SUM_PeakkVA,"")', "SUM_PeakPF", "0.000"),
        ("Peak scenario", "=INDEX(SCN_Names,MATCH(MAX(SCN_S),SCN_S,0))", "SUM_PeakName", None),
        ("Peak transient incl. largest start (kVA, arithmetic add)", "=MAX(SCN_Trans)", "SUM_Trans", "#,##0.0"),
        ("Daily auxiliary energy - normal day (kWh/day)", "=SUM(tblLoad[Day_kWh])", "SUM_DaykWh", "#,##0.0"),
        ("Annual auxiliary energy (MWh/yr) = daily × days per year input", "=SUM_DaykWh*C{days}/1000", "SUM_YrMWh", "#,##0.0"),
        ("Days per year for annualisation (input)", None, "SUM_Days", "0"),
        ("Auxiliary energy per BESS MWh (kWh/day per MWh)", '=IF(N(DB_BESS_MWh)>0,SUM_DaykWh/DB_BESS_MWh,"")', "SUM_kWhPerMWh", "0.00"),
        ("Peak aux demand as % of BESS rated power", '=IF(N(DB_BESS_MW)>0,SUM_PeakkW/(DB_BESS_MW*1000)*100,"")', "SUM_PctMW", "0.00"),
        ("DC loads listed (kW, informative, not in AC totals)", "=SUM(tblLoad[DC_kW])", "SUM_DCkW", "#,##0.00"),
        ("Load rows valid / active", '=SUM(tblLoad[Valid])&" / "&SUM(tblLoad[Active])', "SUM_Rows", None),
        ("Load rows with errors", '=COUNTIF(tblLoad[Check],"ERR*")', "SUM_ErrRows", "0"),
        ("Non-linear share of demand kVA (site)", '=IF(SUM(tblLoad[Run_kVA])>0,SUMIFS(tblLoad[Run_kVA],tblLoad[NL_e],"Y")/SUM(tblLoad[Run_kVA]),0)', "SUM_NL", "0.0%"),
    ]
    r = 4
    days_row = None
    for i, (lab, f, nm, fmt) in enumerate(kp):
        if nm == "SUM_Days":
            days_row = 4 + i
    for lab, f, nm, fmt in kp:
        ws.cell(row=r, column=1, value=lab).border = BRD
        c = ws.cell(row=r, column=2)
        if f is None:
            c.value = cfg.get("days", 365)
            c.fill = F_IN
            c.font = FN_IN
        else:
            c.value = f"=SUM_DaykWh*B{days_row}/1000" if "{days}" in f else f
            c.fill = F_CALC
        if fmt:
            c.number_format = fmt
        c.border = BRD
        add_name(wb, nm, f"Summary!$B${r}")
        r += 1
    # scenario comparison
    r += 1
    ws.cell(row=r, column=1, value="Operating scenario comparison").font = FN_B
    r += 1
    for j, h in enumerate(["Scenario", "kW", "kVA", "PF", "% of peak", "Transient kVA"]):
        c = ws.cell(row=r, column=1 + j, value=h)
        c.fill = F_HDR
        c.font = FN_HDR
    sc0 = r + 1
    for i in range(7):
        rr = sc0 + i
        ws.cell(row=rr, column=1, value=f"=INDEX(SCN_Names,{i + 1})")
        ws.cell(row=rr, column=2, value=f"=INDEX(SCN_P,{i + 1})").number_format = "#,##0.0"
        ws.cell(row=rr, column=3, value=f"=INDEX(SCN_S,{i + 1})").number_format = "#,##0.0"
        ws.cell(row=rr, column=4, value=f'=IF(C{rr}>0,B{rr}/C{rr},"")').number_format = "0.000"
        ws.cell(row=rr, column=5, value=f'=IF(MAX(SCN_S)>0,C{rr}/MAX(SCN_S),0)').number_format = "0.0%"
        ws.cell(row=rr, column=6, value=f"=INDEX(SCN_Trans,{i + 1})").number_format = "#,##0.0"
        for j in range(1, 7):
            ws.cell(row=rr, column=j).fill = F_CALC
    add_name(wb, "SUM_ScenTbl", f"Summary!$A${sc0}:$C${sc0 + 6}")
    r = sc0 + 8
    # category breakdown
    ws.cell(row=r, column=1, value="Load category breakdown").font = FN_B
    ws.cell(row=r, column=3, value="Breakdown scenario:")
    bc = ws.cell(row=r, column=4, value=cfg.get("brk_scen", "MAXSIM"))
    bc.fill = F_IN
    bc.font = FN_IN
    d = DataValidation(type="list", formula1="=lstScenCodes")
    ws.add_data_validation(d)
    d.add(f"D{r}")
    brk = f"$D${r}"
    r += 1
    for j, h in enumerate(["Category", "Connected kW", "Connected kVA", "Scenario demand kW", "Share of scenario kW", "Daily kWh", "Rows"]):
        c = ws.cell(row=r, column=1 + j, value=h)
        c.fill = F_HDR
        c.font = FN_HDR
    cat0 = r + 1
    pcols = ",".join(f"tblLoad[P_{s}]" for s in SC)
    for k in range(1, 31):
        rr = cat0 + k - 1
        ws.cell(row=rr, column=1, value=f'=IF(ROWS(tblCat[Category])>={k},IF(INDEX(tblCat[Category],{k})="","",INDEX(tblCat[Category],{k})),"")')
        ws.cell(row=rr, column=2, value=f'=IF($A{rr}="","",SUMIFS(tblLoad[Conn_kW],tblLoad[Cat_e],$A{rr}))')
        ws.cell(row=rr, column=3, value=f'=IF($A{rr}="","",SUMIFS(tblLoad[Conn_kVA],tblLoad[Cat_e],$A{rr}))')
        ws.cell(row=rr, column=4, value=f'=IF($A{rr}="","",SUMIFS(CHOOSE(MATCH({brk},lstScenCodes,0),{pcols}),tblLoad[Cat_e],$A{rr}))')
        ws.cell(row=rr, column=5, value=f'=IF($A{rr}="","",IF(INDEX(SCN_P,MATCH({brk},lstScenCodes,0))>0,D{rr}/INDEX(SCN_P,MATCH({brk},lstScenCodes,0)),0))')
        ws.cell(row=rr, column=6, value=f'=IF($A{rr}="","",SUMIFS(tblLoad[Day_kWh],tblLoad[Cat_e],$A{rr}))')
        ws.cell(row=rr, column=7, value=f'=IF($A{rr}="","",COUNTIFS(tblLoad[Cat_e],$A{rr},tblLoad[Valid],1))')
        for j in range(1, 8):
            c = ws.cell(row=rr, column=j)
            c.fill = F_CALC
            c.number_format = "0.0%" if j == 5 else "#,##0.0"
    rr = cat0 + 30
    ws.cell(row=rr, column=1, value="Uncategorised / not in tblCat")
    ws.cell(row=rr, column=2, value=f"=SUM(tblLoad[Conn_kW])-SUM(B{cat0}:B{rr - 1})")
    ws.cell(row=rr, column=3, value=f"=SUM(tblLoad[Conn_kVA])-SUM(C{cat0}:C{rr - 1})")
    ws.cell(row=rr, column=4, value=f"=INDEX(SCN_P,MATCH({brk},lstScenCodes,0))-SUM(D{cat0}:D{rr - 1})")
    ws.cell(row=rr, column=6, value=f"=SUM(tblLoad[Day_kWh])-SUM(F{cat0}:F{rr - 1})")
    for j in (2, 3, 4, 6):
        ws.cell(row=rr, column=j).number_format = "#,##0.0"
        ws.cell(row=rr, column=j).fill = F_CALC
    add_name(wb, "SUM_CatNames", f"Summary!$A${cat0}:$A${cat0 + 29}")
    add_name(wb, "SUM_CatConn", f"Summary!$B${cat0}:$B${cat0 + 29}")
    # TX allocation
    r = rr + 3
    ws.cell(row=r, column=1, value="Centralized / distributed load allocation (per transformer)").font = FN_B
    ws.cell(row=r, column=4, value='=IF(COUNTIF(tblTX[TX_ID],"?*")>20,"WARNING: more than 20 transformers - list truncated (calculations unaffected, see sheet 6)","")').font = Font(bold=True, color="C00000")
    r += 1
    hd = ["TX_ID", "Role", "Connected kVA", "Design kVA", "Gov. scenario", "Evaluated rating kVA", "Loading %", "Rating status"]
    for j, h in enumerate(hd):
        c = ws.cell(row=r, column=1 + j, value=h)
        c.fill = F_HDR
        c.font = FN_HDR
    tx0 = r + 1
    for k in range(1, 21):
        rr = tx0 + k - 1
        ws.cell(row=rr, column=1, value=f'=IF(ROWS(tblTX[TX_ID])>={k},IF(INDEX(tblTX[TX_ID],{k})="","",INDEX(tblTX[TX_ID],{k})),"")')
        for j, col in enumerate(["Role", "Conn_kVA", "Design_kVA", "Gov", "Eval_kVA", "Load_pct", "RatingStatus"]):
            c = ws.cell(row=rr, column=2 + j, value=f'=IF($A{rr}="","",INDEX(tblTX[{col}],{k})&"")' if col in ("Role", "Gov", "RatingStatus")
                        else f'=IF($A{rr}="","",INDEX(tblTX[{col}],{k}))')
            c.fill = F_CALC
            c.number_format = "#,##0.0"
        ws.cell(row=rr, column=1).fill = F_CALC
    add_name(wb, "SUM_TXid", f"Summary!$A${tx0}:$A${tx0 + 19}")
    add_name(wb, "SUM_TXload", f"Summary!$G${tx0}:$G${tx0 + 19}")
    rr = tx0 + 21
    ws.cell(row=rr, column=1, value="Site coincident demand (kVA)")
    ws.cell(row=rr, column=2, value="=CMP_SiteDesign").number_format = "#,##0.0"
    ws.cell(row=rr + 1, column=1, value="Σ per-transformer design demand (kVA)")
    ws.cell(row=rr + 1, column=2, value="=CMP_SumTX").number_format = "#,##0.0"
    ws.cell(row=rr + 2, column=1, value="Centralized single-TX calculated minimum / recommended (kVA)")
    ws.cell(row=rr + 2, column=2, value="=CMP_CentMin").number_format = "#,##0.0"
    ws.cell(row=rr + 2, column=3, value="=CMP_CentRec").number_format = "#,##0.#"
    ws.cell(row=rr + 3, column=1, value="Distributed: Σ evaluated ratings (kVA)")
    ws.cell(row=rr + 3, column=2, value="=CMP_DistInst").number_format = "#,##0.0"
    for i in range(4):
        ws.cell(row=rr + i, column=2).fill = F_CALC
    ws.cell(row=rr + 2, column=3).fill = F_CALC
    return dict(sc0=sc0, cat0=cat0, tx0=tx0)


# ======================================================================= VALIDATION & STANDARDS
STANDARDS = [
    # id, standard, title, edition (verified 2026-11-02 via publisher/distributor listings unless noted), status, juris, applicability
    ("S01", "IEC 60076-1", "Power transformers - Part 1: General", "2011 (Ed. 3.0); revision in progress (prEN IEC 60076-1) - check before use",
     "Mandatory where specified by contract / national adoption", "IEC", "Rating definitions, normal service conditions, tolerances (sheet 6)"),
    ("S02", "IEC 60076-2", "Power transformers - Part 2: Temperature rise for liquid-immersed transformers", "Edition NOT verified in this build",
     "Contract / national adoption", "IEC", "Temperature-rise & altitude/ambient corrections for liquid-filled units"),
    ("S03", "IEC 60076-5", "Power transformers - Part 5: Ability to withstand short circuit", "Edition NOT verified", "Contract", "IEC", "Impedance / through-fault"),
    ("S04", "IEC 60076-7", "Power transformers - Part 7: Loading guide for mineral-oil-immersed transformers", "2018 (to be confirmed by user)",
     "Guide (recommended practice)", "IEC", "Overload / contingency loading > 100 %"),
    ("S05", "IEC 60076-11", "Power transformers - Part 11: Dry-type transformers", "2018 (Ed. 2.0) - verified", "Contract / national adoption", "IEC", "Dry-type rating & service conditions"),
    ("S06", "IEC 60076-12", "Power transformers - Part 12: Loading guide for dry-type power transformers", "Edition NOT verified", "Guide", "IEC", "Dry-type overload"),
    ("S07", "IEC 60364 series (incl. -4-41, -4-43, -5-52, -5-54)", "Low-voltage electrical installations",
     "60364-5-52: 2009 + AMD1:2024 (consolidated) - verified; other parts NOT verified", "Mandatory where adopted nationally", "IEC",
     "LV feeder current-carrying capacity, overload coordination, voltage drop, earthing (sheet 7)"),
    ("S08", "IEC 60909-0", "Short-circuit currents in three-phase a.c. systems - Part 0: Calculation of currents", "2016 (Ed. 2.0) - verified",
     "Calculation method (contract)", "IEC", "Voltage factor c; prospective LV fault level (sheet 6, preliminary)"),
    ("S09", "IEC 61439-1 / -2", "Low-voltage switchgear and controlgear assemblies", "Edition NOT verified", "Product standard", "IEC", "LV board ratings (outside workbook scope)"),
    ("S10", "IEC 60947-2", "LV switchgear - Circuit-breakers", "Edition NOT verified", "Product standard", "IEC", "Device I2 ratio, Icu (user inputs sheet 7)"),
    ("S11", "IEC 62040-3", "Uninterruptible power systems - performance", "Edition NOT verified", "Product standard", "IEC", "UPS input rating & recharge"),
    ("S12", "IEC 62933-5-2", "Electrical energy storage systems - Part 5-2: Safety requirements for grid-integrated EES systems - Electrochemical based systems",
     "2025 (Ed. 2.0) - verified", "Contract / national adoption", "IEC", "Safety-related auxiliaries (ventilation, detection, essential supplies)"),
    ("S13", "IEC 61936-1", "Power installations exceeding 1 kV a.c. - Common rules", "Edition NOT verified", "National adoption", "IEC", "MV side of aux transformers"),
    ("S14", "IEEE C57.12.00", "General Requirements for Liquid-Immersed Distribution, Power and Regulating Transformers", "2021 - verified",
     "Mandatory where specified (US practice)", "US", "Ratings, service conditions, altitude/ambient (sheet 6)"),
    ("S15", "IEEE C57.12.01", "General Requirements for Dry-Type Distribution and Power Transformers", "Edition NOT verified", "US practice", "US", "Dry-type ratings & service conditions"),
    ("S16", "IEEE C57.91", "Guide for Loading Mineral-Oil-Immersed Transformers and Step-Voltage Regulators", "2025 - verified (supersedes 2011)",
     "Guide", "US", "Overload / contingency loading"),
    ("S17", "IEEE C57.96", "Guide for Loading Dry-Type Distribution and Power Transformers", "Edition NOT verified", "Guide", "US", "Dry-type overload"),
    ("S18", "IEEE C57.110", "Recommended Practice for Establishing Liquid-Immersed and Dry-Type Transformer Capability When Supplying Nonsinusoidal Load Currents",
     "2018 - verified", "Recommended practice", "Both", "Harmonic derating factor Kharm (sheet 6)"),
    ("S19", "IEEE 519", "Standard for Harmonic Control in Electric Power Systems", "2022 - verified", "Standard (PCC limits; contract)", "Both", "Harmonic limits at PCC (aux system contribution)"),
    ("S20", "IEEE 80", "Guide for Safety in AC Substation Grounding", "2013 - verified current; revision PAR active", "Guide", "Both", "Earthing of aux transformer / MV yard (outside workbook)"),
    ("S21", "IEEE 142", "Recommended Practice for Grounding of Industrial and Commercial Power Systems", "Edition NOT verified", "Recommended practice", "US", "LV system earthing"),
    ("S22", "IEEE 1584", "Guide for Performing Arc-Flash Hazard Calculations", "Edition NOT verified", "Guide", "US", "Arc-flash (requires full SC study)"),
    ("S23", "IEEE 946 / IEEE 485", "DC auxiliary power systems for generating stations / Lead-acid battery sizing", "Editions NOT verified", "Recommended practice", "Both", "DC system & charger sizing"),
    ("S24", "IEEE 2030.2.1", "Guide for Design, Operation and Maintenance of BESS integrated with power systems", "Edition NOT verified", "Guide", "Both", "BESS auxiliary philosophy"),
    ("S25", "NFPA 70 (NEC)", "National Electrical Code - incl. Art. 210/215 (circuits/feeders), 240 (overcurrent), 310 (conductors), 450 (transformers), 480 (batteries), 695 (fire pumps), 700/701/702 (emergency/standby), 706 (ESS)",
     "2026 edition - verified published; adoption varies by state/AHJ", "Mandatory where adopted by AHJ", "US", "Feeder/OCPD sizing (ContMult), transformer protection, fire-pump supply, ESS"),
    ("S26", "NFPA 855", "Standard for the Installation of Stationary Energy Storage Systems", "2026 edition - verified", "Mandatory where adopted by AHJ", "US",
     "Safety systems requiring power (ventilation, detection, explosion control) and their supply reliability"),
    ("S27", "UL 9540 / UL 9540A", "Energy Storage Systems and Equipment / Test method for thermal runaway fire propagation", "UL 9540: 2023 (3rd ed.) - verified; 9540A edition NOT verified",
     "Product listing (US)", "US", "Equipment listing; informs auxiliary safety loads"),
    ("S28", "NFPA 20 / NFPA 72", "Fire pumps / National Fire Alarm and Signaling Code", "Editions NOT verified", "Mandatory where adopted", "US", "Fire-pump supply & fire-alarm power"),
    ("S29", "Local grid code / utility & client specifications", "Project-specific requirements", "Project specific", "Project-specific", "Both",
     "Margins, loading limits, redundancy, voltage criteria (Design Basis)"),
]


def build_validation(wb, ws, cfg):
    title(ws, "Sheet 8 - Engineering Validation & Compliance Register",
          "Results are formula-driven from the model. 'Clause' entries are given only where known; ALL clauses and acceptance "
          "criteria must be verified against licensed copies of the standards. The workbook does not claim compliance.")
    hd = ["Ref", "Design requirement", "Applicable standard", "Edition / year", "Clause / section", "Applicability",
          "Calculation / design check", "Acceptance criterion", "Result", "Evidence / reference", "Engineering remarks"]
    widths = [6, 40, 22, 16, 18, 14, 40, 34, 16, 26, 36]
    for j, (h, w) in enumerate(zip(hd, widths)):
        c = ws.cell(row=4, column=1 + j, value=h)
        c.fill = F_HDR
        c.font = FN_HDR
        c.alignment = WRAP
        ws.column_dimensions[CL(1 + j)].width = w
    items = [
        ("Design basis mandatory inputs complete", "Project specification", "-", "-", "All",
         "Count of ERR flags on Design Basis", "Zero errors", '=IF(DB_ErrCount=0,"PASS","FAIL")', "Design Basis col F"),
        ("Design basis criteria entered (margins, limits, derating, c-factor)", "Project specification", "-", "-", "All",
         "Count of REVIEW flags on Design Basis", "Zero open review items", '=IF(DB_RevCount=0,"PASS","REVIEW")', "Design Basis col F"),
        ("Load schedule data integrity", "Engineering practice", "-", "-", "All",
         "Rows with ERR in tblLoad[Check]", "Zero error rows", '=IF(COUNTIF(tblLoad[Check],"ERR*")=0,IF(SUM(tblLoad[Valid])>0,"PASS","REVIEW"),"FAIL")', "Load Schedule Check column"),
        ("Load schedule warnings resolved", "Engineering practice", "-", "-", "All",
         "Rows with WARN", "Zero warnings", '=IF(COUNTIF(tblLoad[Check],"WARN*")=0,"PASS","REVIEW")', "Load Schedule Check column"),
        ("Equipment ratings traceable to datasheets", "Engineering practice / QA", "-", "-", "All",
         "Valid rows without Reference", "All rated rows referenced", '=IF(COUNTIFS(tblLoad[Valid],1,tblLoad[Reference],"")=0,"PASS","REVIEW")', "Load Schedule Reference column"),
        ("Mutually exclusive loads not double-counted", "Engineering practice", "-", "-", "All",
         "Σ counted ≤ Σ raw in every scenario; one member per ExclGroup", "Consistency PASS", "=SCN_Consistency", "Scenario Analysis"),
        ("Scenario profiles reviewed for project", "Engineering practice", "-", "-", "All",
         "Design Basis confirmation", "Confirmed 'Yes'", '=IF(DB_ProfRev="Yes","PASS","REVIEW")', "Scenario Analysis tblProf"),
        ("Max-simultaneous scenario envelopes normal & max-cooling", "Engineering practice", "-", "-", "All",
         "S(MAXSIM) ≥ S(NORM), S(MAXCOOL)", "True", "=SCN_MaxSimChk", "Scenario Analysis"),
        ("All AC loads assigned to a transformer", "Engineering practice", "-", "-", "All",
         "Valid AC rows with blank TX", "Zero", '=IF(COUNTIFS(tblLoad[IsAC],1,tblLoad[TX],"")=0,"PASS","REVIEW")', "Load Schedule TX column"),
        ("Transformer definition errors", "IEC 60076-1 / IEEE C57.12.00", "2011 / 2021", "Not verified", "All",
         "tblTX Check = ERR", "Zero", '=IF(COUNTIF(tblTX[Check],"ERR*")=0,"PASS","FAIL")', "Transformer Sizing Check"),
        ("Transformer loading within permissible (normal)", "IEC 60076-1 / IEEE C57.12.00; client spec", "2011 / 2021", "Not verified", "All",
         "Design kVA ÷ derated capacity", "≤ DB_MaxLoad", '=IF(COUNTIF(tblTX[TX_ID],"?*")=0,"REVIEW",IF(COUNTIFS(tblTX[TX_ID],"?*",tblTX[LoadDer_pct],">"&DB_MaxLoad*100+0.000001)>0,"FAIL",IF(COUNTIFS(tblTX[TX_ID],"?*",tblTX[DerateOK],0)>0,"REVIEW","PASS")))', "tblTX[LoadDer_pct]"),
        ("Redundancy (N+1 / 2N) contingency capacity", "Client spec; IEC 60076-7 / IEEE C57.91 if > 100 %", "2018 / 2025", "Not verified", "Where N+1/2N",
         "Group demand ÷ units in service ÷ derated capacity", "≤ DB_MaxLoadCont", '=IF(COUNTIFS(tblTX[Scheme_e],"<>N",tblTX[TX_ID],"?*")=0,"NOT APPLICABLE",IF(COUNTIFS(tblTX[ContLoad_pct],">"&DB_MaxLoadCont*100+0.000001)=0,"PASS","FAIL"))', "tblTX[ContLoad_pct]"),
        ("Transformer rating formally selected (not provisional)", "QA / design review", "-", "-", "All",
         "Selected_kVA entered and ≥ calculated minimum", "All selected & adequate", '=IF(COUNTIF(tblTX[RatingStatus],"FAIL*")>0,"FAIL",IF(COUNTIF(tblTX[RatingStatus],"SELECTED*")=COUNTIF(tblTX[TX_ID],"?*"),IF(COUNTIF(tblTX[TX_ID],"?*")>0,"PASS","REVIEW"),"REVIEW"))', "tblTX[RatingStatus]"),
        ("Ambient & altitude within rating basis or derated", "IEC 60076-1/-2/-11; IEEE C57.12.00/.01", "see matrix", "Not verified", "All",
         "Site ambient/altitude vs rating basis; derating coefficient entered", "Within basis or derating entered", '=IF(COUNTIF(tblTX[Check],"REVIEW: ambient*")+COUNTIF(tblTX[Check],"REVIEW: altitude*")=0,"PASS","REVIEW")', "tblTX[Kamb],[Kalt]"),
        ("Harmonic / non-linear load capability", "IEEE C57.110; IEEE 519", "2018 / 2022", "Not verified", "Where NL share > threshold",
         "NL share vs threshold, Kharm entered", "Kharm entered where required", '=IF(NOT(ISNUMBER(DB_NLThr)),"REVIEW",IF(COUNTIF(tblTX[Check],"REVIEW: non-linear*")=0,"PASS","REVIEW"))', "tblTX[NL_share]"),
        ("Transformer voltage regulation", "Project criterion", "-", "-", "All",
         "ε at design load", "≤ DB_RegLim", '=IF(NOT(ISNUMBER(DB_RegLim)),"REVIEW",IF(COUNTIFS(tblTX[TX_ID],"?*",tblTX[Reg_pct],"")>0,"REVIEW",IF(COUNTIFS(tblTX[Reg_pct],">"&DB_RegLim*100)=0,"PASS","FAIL")))', "tblTX[Reg_pct]"),
        ("Motor-starting voltage dip", "Project / equipment criterion", "-", "-", "Where motors",
         "Largest start kVA ÷ (start + bus Ssc)", "≤ DB_DipLim", '=IF(NOT(ISNUMBER(DB_DipLim)),"REVIEW",IF(COUNTIFS(tblTX[Dip_pct],">"&DB_DipLim*100)=0,IF(COUNTIFS(tblTX[TX_ID],"?*",tblTX[Dip_pct],"")>0,"REVIEW","PASS"),"FAIL"))', "tblTX[Dip_pct]"),
        ("Transformer energisation inrush data obtained", "Protection coordination practice", "-", "-", "All",
         "InrushMult entered", "All TX", '=IF(COUNTIFS(tblTX[TX_ID],"?*",tblTX[InrushMult],"")=0,"PASS","REVIEW")', "tblTX[Inrush_A]"),
        ("Feeder overload coordination Ib ≤ In ≤ Iz", "IEC 60364-4-43 / NEC 210, 215, 240", "see matrix", "IEC 60364-4-43 cl. 433.1 (verify)", "All feeders",
         "Chk_In, Chk_Iz", "PASS all", '=IF(COUNTIF(tblFdr[FeederID],"?*")=0,"REVIEW",IF(COUNTIF(tblFdr[Chk_In],"FAIL")+COUNTIF(tblFdr[Chk_Iz],"FAIL")>0,"FAIL",IF(COUNTIF(tblFdr[Chk_In],"REVIEW*")+COUNTIF(tblFdr[Chk_Iz],"REVIEW*")>0,"REVIEW","PASS")))', "tblFdr"),
        ("Feeder I2 ≤ 1.45 Iz", "IEC 60364-4-43", "see matrix", "cl. 433.1 (verify)", "IEC installations",
         "Chk_I2", "PASS all", '=IF(DB_Jurisdiction="US (NEC/NFPA/IEEE)","NOT APPLICABLE",' + 'IF(COUNTIF(tblFdr[Chk_I2],"FAIL")>0,"FAIL",IF(COUNTIF(tblFdr[Chk_I2],"REVIEW*")>0,"REVIEW","PASS")))', "tblFdr[Chk_I2]"),
        ("Feeder voltage drop", "IEC 60364-5-52 / NEC (informational notes)", "2009+A1:2024 / 2026", "Not verified", "All feeders",
         "ΔV%", "≤ DB_VDLim", '=IF(COUNTIF(tblFdr[Chk_VD],"FAIL")>0,"FAIL",IF(COUNTIF(tblFdr[Chk_VD],"REVIEW*")>0,"REVIEW",IF(COUNTIF(tblFdr[FeederID],"?*")=0,"REVIEW","PASS")))', "tblFdr[VD_pct]"),
        ("Device breaking capacity ≥ prospective fault", "IEC 60909-0 / IEC 60947-2 / NEC 110.9", "2016 / - / 2026", "Not verified", "All feeders",
         "Icu vs Isc at TX LV terminals", "Icu ≥ Isc", '=IF(COUNTIF(tblFdr[Chk_Icu],"FAIL")>0,"FAIL",IF(COUNTIF(tblFdr[Chk_Icu],"REVIEW*")>0,"REVIEW",IF(COUNTIF(tblFdr[FeederID],"?*")=0,"REVIEW","PASS")))', "tblFdr[Chk_Icu]"),
        ("LV phase balance", "Project criterion", "-", "-", "Where 1PH loads",
         "Imbalance % per TX", "≤ DB_ImbLim", '=IF(COUNTIF(tblTX[BalCheck],"FAIL")>0,"FAIL",IF(COUNTIF(tblTX[BalCheck],"REVIEW*")>0,"REVIEW",IF(COUNTIF(tblTX[TX_ID],"?*")=0,"REVIEW","PASS")))', "tblTX[BalCheck]"),
        ("Motor-starting / inrush / restoration study (screening only in workbook)", "Engineering practice; NFPA 20 where fire pumps", "-", "-", "Where starting data entered",
         "Rows with starting multiplier > 1", "Detailed study or manufacturer confirmation", '=IF(COUNTIF(tblLoad[StartDelta],">0")=0,"NOT APPLICABLE","REVIEW")', "tblLoad[Start_kVA]"),
        ("Reporting blocks cover all defined items", "Workbook integrity", "-", "-", "All",
         "Transformers ≤ 20 (sheets 4, 5, 9) and categories ≤ 30 (sheet 5); calculations themselves are unlimited", "Within display capacity",
         '=IF(AND(COUNTIF(tblTX[TX_ID],"?*")<=20,COUNTIF(tblCat[Category],"?*")<=30),"PASS","FAIL: display blocks truncated - extend slots")', "Sheets 4, 5, 9 slot lists"),
        ("Fire pump supply arrangement", "NFPA 20 / NEC 695 (US); local fire code", "NOT verified / 2026", "Not verified", "Where fire pumps",
         "Manual review of supply source, locked-rotor capability", "Per AHJ / code", '=IF(SUMIFS(tblLoad[Conn_kW],tblLoad[Cat_e],"FIRE-PUMP")=0,"NOT APPLICABLE","REVIEW")', "Load rows category FIRE-PUMP"),
        ("ESS safety-system power supply reliability", "NFPA 855 / IEC 62933-5-2; NEC 706", "2026 / 2025 / 2026", "Not verified", "All BESS",
         "Manual review: ventilation, detection, suppression supplies & backup", "Per AHJ / code", '="REVIEW"', "SupplyClass column"),
        ("Standards editions & clauses verified", "All", "-", "-", "All",
         "Design Basis confirmation", "Confirmed 'Yes'", '=IF(DB_StdRev="Yes","PASS","REVIEW")', "Standards matrix below"),
    ]
    global N_VAL
    N_VAL = len(items)
    r = 5
    for i, it in enumerate(items):
        row = [f"V{i + 1:02d}", it[0], it[1], it[2], it[3], it[4], it[5], it[6], it[7], it[8], ""]
        for j, v in enumerate(row):
            c = ws.cell(row=r, column=1 + j, value=v)
            c.alignment = WRAP
            c.border = BRD
            if j == 8:
                c.fill = F_CALC
            if j == 10:
                c.fill = F_IN
                c.font = FN_IN
        r += 1
    status_cf(ws, f"I5:I{r - 1}")
    add_name(wb, "VAL_Results", f"'Validation'!$I$5:$I${r - 1}")
    ws.cell(row=r + 1, column=1, value="Register summary").font = FN_B
    for k, lab in enumerate(["PASS", "FAIL", "REVIEW", "NOT APPLICABLE"]):
        ws.cell(row=r + 2 + k, column=2, value=lab)
        c = ws.cell(row=r + 2 + k, column=3, value=f'=COUNTIF(VAL_Results,"{lab}*")')
        c.fill = F_CALC
        add_name(wb, f"VAL_N_{lab.replace(' ', '_')}", f"'Validation'!$C${r + 2 + k}")
    # standards matrix
    r = r + 8
    ws.cell(row=r, column=1, value="Standards applicability matrix (editions checked against publisher / distributor listings on 2026-11-02 where marked 'verified')").font = FN_B
    hd2 = ["Ref", "Standard", "Title (abridged)", "Edition / status", "Requirement type", "Jurisdiction",
           "Applicability to this workbook", "Applies to project?", "Full text needed for final verification", "Notes"]
    for j, h in enumerate(hd2):
        c = ws.cell(row=r + 1, column=1 + j, value=h)
        c.fill = F_HDR
        c.font = FN_HDR
        c.alignment = WRAP
    for i, s in enumerate(STANDARDS):
        rr = r + 2 + i
        row = list(s)
        juris = s[5]
        app = ('=IF(DB_Jurisdiction="","REVIEW: set jurisdiction",' +
               ('"YES (if specified)")' if juris == "Both" else
                f'IF(OR(DB_Jurisdiction="Mixed / Other",DB_Jurisdiction="{"IEC" if juris == "IEC" else "US (NEC/NFPA/IEEE)"}"),"YES (if specified)","NOT APPLICABLE (check)"))'))
        vals = row + [app, "Yes", ""]
        for j, v in enumerate(vals):
            c = ws.cell(row=rr, column=1 + j, value=v)
            c.alignment = WRAP
            c.border = BRD
            if j == 7:
                c.fill = F_CALC
            if j == 9:
                c.fill = F_IN
                c.font = FN_IN
    ws.cell(row=r + 3 + len(STANDARDS), column=1, value=(
        "Mandatory = legally required only where adopted by the authority having jurisdiction or invoked by contract. "
        "Guides / recommended practices are good engineering practice. No standard text is reproduced here.")).font = FN_NOTE
    ws.freeze_panes = "C5"


# ======================================================================= DASHBOARD
def build_dashboard(wb, ws, cfg):
    title(ws, "Sheet 9 - Engineering Dashboard", "All values linked to calculation sheets. Charts update automatically.")
    ws.column_dimensions["A"].width = 3
    for c in "BCDEFGHIJKLMN":
        ws.column_dimensions[c].width = 14
    tiles = [
        ("Connected load (kW)", "=SUM_ConnkW", "#,##0"),
        ("Connected load (kVA)", "=SUM_ConnkVA", "#,##0"),
        ("Maximum demand (kVA)", "=SUM_PeakkVA", "#,##0"),
        ("Maximum demand (kW)", "=SUM_PeakkW", "#,##0"),
        ("Peak scenario", "=SUM_PeakName", None),
        ("Daily aux energy (kWh)", "=SUM_DaykWh", "#,##0"),
        ("Σ evaluated TX rating (kVA)", "=SUM(tblTX[Eval_kVA])", "#,##0"),
        ("Highest TX loading (%)", "=MAX(tblTX[Load_pct])", "0.0"),
        ("Min. spare capacity (kVA)", '=IF(COUNTIF(tblTX[TX_ID],"?*")=0,"",MIN(tblTX[Spare_kVA]))', "#,##0"),
        ("TX ratings provisional", '=COUNTIF(tblTX[RatingStatus],"PROVISIONAL*")', "0"),
        ("Validation FAIL", "=VAL_N_FAIL", "0"),
        ("Validation REVIEW", "=VAL_N_REVIEW", "0"),
    ]
    r = 4
    for i, (lab, f, fmt) in enumerate(tiles):
        col = 2 + (i % 4) * 3
        rr = r + (i // 4) * 3
        a = ws.cell(row=rr, column=col, value=lab)
        a.font = Font(size=9, color="404040")
        a.fill = F_KPI
        ws.merge_cells(start_row=rr, start_column=col, end_row=rr, end_column=col + 1)
        b = ws.cell(row=rr + 1, column=col, value=f)
        b.font = Font(size=16, bold=True, color="1F3864")
        b.fill = F_KPI
        b.alignment = Alignment(horizontal="left")
        if fmt:
            b.number_format = fmt
        ws.merge_cells(start_row=rr + 1, start_column=col, end_row=rr + 1, end_column=col + 1)
    # selected transformer rating table
    r = 14
    ws.cell(row=r, column=2, value="Transformers: rating, loading, spare & status").font = FN_B
    hd = ["TX_ID", "Req. min kVA", "Recommended std", "Selected kVA", "Loading %", "Spare kVA", "Contingency %", "Status", "Check"]
    for j, h in enumerate(hd):
        c = ws.cell(row=r + 1, column=2 + j, value=h)
        c.fill = F_HDR
        c.font = FN_HDR
    for k in range(1, 21):
        rr = r + 1 + k
        ws.cell(row=rr, column=2, value=f'=IF(ROWS(tblTX[TX_ID])>={k},IF(INDEX(tblTX[TX_ID],{k})="","",INDEX(tblTX[TX_ID],{k})),"")')
        for j, col in enumerate(["Req_min", "Rec_std", "Selected_kVA", "Load_pct", "Spare_kVA", "ContLoad_pct", "RatingStatus", "Check"]):
            txt = col in ("RatingStatus", "Check", "Rec_std", "Selected_kVA", "ContLoad_pct")
            c = ws.cell(row=rr, column=3 + j, value=f'=IF($B{rr}="","",INDEX(tblTX[{col}],{k}){"&\"\"" if col in ("RatingStatus", "Check") else ""})')
            c.number_format = "#,##0.0"
            c.fill = F_CALC
    status_cf(ws, "I16:J35")
    # outstanding checks
    r = 38
    ws.cell(row=r, column=2, value="Outstanding engineering checks (validation register items not PASS / N/A)").font = FN_B
    for k in range(1, N_VAL + 1):
        rr = r + k
        ws.cell(row=rr, column=2, value=f'=IF(OR(INDEX(VAL_Results,{k})="PASS",INDEX(VAL_Results,{k})="NOT APPLICABLE"),"",INDEX(Validation!$A$5:$A$60,{k})&"  "&INDEX(Validation!$B$5:$B$60,{k}))')
        ws.cell(row=rr, column=9, value=f'=IF($B{rr}="","",INDEX(VAL_Results,{k}))')
    status_cf(ws, f"I{r + 1}:I{r + N_VAL}")
    return ws


def add_charts(wb, ws_dash, ws_sum, pos):
    ch = BarChart()
    ch.type = "col"
    ch.title = "Scenario comparison - apparent power (kVA)"
    ch.y_axis.title = "kVA"
    data = Reference(ws_sum, min_col=3, min_row=pos["sc0"] - 1, max_row=pos["sc0"] + 6)
    cats = Reference(ws_sum, min_col=1, min_row=pos["sc0"], max_row=pos["sc0"] + 6)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    ch.height, ch.width = 7.5, 16
    ch.legend = None
    ws_dash.add_chart(ch, "L14")
    ch2 = BarChart()
    ch2.type = "bar"
    ch2.title = "Connected load by category (kW)"
    data = Reference(ws_sum, min_col=2, min_row=pos["cat0"] - 1, max_row=pos["cat0"] + 15)
    cats = Reference(ws_sum, min_col=1, min_row=pos["cat0"], max_row=pos["cat0"] + 15)
    ch2.add_data(data, titles_from_data=True)
    ch2.set_categories(cats)
    ch2.height, ch2.width = 9, 16
    ch2.legend = None
    ws_dash.add_chart(ch2, "L38")
    ch3 = BarChart()
    ch3.type = "col"
    ch3.title = "Transformer loading (% of evaluated rating)"
    data = Reference(ws_sum, min_col=7, min_row=pos["tx0"] - 1, max_row=pos["tx0"] + 11)
    cats = Reference(ws_sum, min_col=1, min_row=pos["tx0"], max_row=pos["tx0"] + 11)
    ch3.add_data(data, titles_from_data=True)
    ch3.set_categories(cats)
    ch3.height, ch3.width = 7.5, 16
    ch3.legend = None
    ws_dash.add_chart(ch3, "B72")


# ======================================================================= METHODOLOGY
METHOD = [
    ("H", "1. Purpose and scope"),
    ("T", "Project-independent tool for preliminary estimation of BESS auxiliary loads, operating-scenario demand, auxiliary transformer sizing "
          "(centralized or distributed, N / N+1 / 2N) and preliminary LV distribution checks. It is NOT a substitute for detailed load-flow, "
          "short-circuit, protection-coordination, arc-flash, harmonic or thermal studies, and does not claim compliance with any standard."),
    ("H", "2. Engineering workflow"),
    ("T", "1) Sheet 1: enter design basis (frequency, voltages, ambient, altitude, margins, criteria). Red/amber flags in column F show missing items."),
    ("T", "2) Sheet 2: review categories and library classes; enter ratings ONLY from datasheets with a source reference (ratings are intentionally blank)."),
    ("T", "3) Sheet 6: define each auxiliary transformer (or supply group) - ID, voltages, type, cooling, impedance, redundancy group & scheme."),
    ("T", "4) Sheet 7: define LV feeders (optional for load estimation; required for distribution checks)."),
    ("T", "5) Sheet 3: enter loads - one row per equipment type per supply point; select LibID to inherit library data or type values directly."),
    ("T", "6) Sheet 4: review scenario definitions and participation profiles; assign mutually exclusive loads a common ExclGroup."),
    ("T", "7) Review sheets 5, 6, 7 results; enter Selected_kVA only after manufacturer / client review; close items in sheet 8; Dashboard (sheet 9)."),
    ("H", "3. Cell conventions"),
    ("T", "Yellow fill / blue text = input. Grey fill = formula (do not overwrite). Tables (tblLoad, tblTX, tblFdr, tblLib, tblCat, tblProf, tblScen) "
          "expand automatically when typing in the row directly below; calculated columns fill down. Inputs are workbook Names DB_<key> (Formulas ▸ Name Manager)."),
    ("H", "4. Load-row equations (sheet 3)"),
    ("T", "Effective value = row input if entered, else library value (LibID), else design-basis default (DF, duty, hours only - flagged INFO)."),
    ("T", "Applied voltage V: 3PH → V_LL of the supplying transformer (or DB_VLL); 1PH L-N → V_LN; 1PH L-L → V_LL; DC → DB_VDC; VOverride overrides."),
    ("T", "Per-unit apparent power S (kVA): RatingType kVA: S = rating; A: S = √3·V·I/1000 (3PH) or V·I/1000 (1PH, DC); kW_in: S = P/PF; "
          "kW_out (mechanical/shaft output): S = Pout/(η·PF). Per-unit P = S·PF; I = S·1000/(√3·V) (3PH) or S·1000/V (1PH/DC)."),
    ("T", "Connected kW = QtyInst·P (all installed units incl. standby). Demand kW = QtyDuty·P·DF (DF = demand/load factor = absorbed ÷ rated input)."),
    ("T", "Scenario raw kW (R_s) = Demand kW × profile factor(s). Counted kW (P_s): if ExclGroup blank, P_s = R_s; otherwise only the first row having the "
          "group maximum apparent power (raw kVA RS_s) is counted with its own P and Q (others 0) - prevents mutually exclusive loads (e.g. heater vs chiller, duty vs standby pump) being summed. "
          "All members of an ExclGroup must be on the same TX and Feeder (enforced as ERR), so site, transformer and feeder totals are consistent."),
    ("T", "Reactive power Q = P·tanφ; tanφ is negative for rows marked Reactive = Lead (capacitive), so leading loads offset lagging kvar. Aggregated apparent power S = √(ΣP² + ΣQ²) (vector sum). Arithmetic ΣS shown for comparison (conservative)."),
    ("T", "Daily energy kWh = Demand kW × duty-cycle factor × hours/day (normal operating day). Annual = daily × days/year input. DC rows are listed for DC-system "
          "information only and excluded from AC totals - represent their AC demand through the charger/UPS row to avoid double counting."),
    ("T", "Starting: Start kVA (one unit) = S·StartMult (blank = 1, i.e. no inrush data). Transient scenario S = S_s + largest (Start kVA − running kVA) of loads "
          "active in that scenario, added arithmetically (conservative). Steady-state demand is never inflated by starting values."),
    ("H", "5. Transformer sizing (sheet 6)"),
    ("T", "Scenario S per TX = √(ΣP_s² + ΣQ_s²) of rows assigned to that TX. Design demand = selected scenario or ENVELOPE (max of all scenarios)."),
    ("T", "Derating K = Kamb·Kalt·Kharm. Kamb = 1 − (Tamb − Tref)·k_T/100 when Tamb > Tref and k_T entered; Kalt = 1 − (Alt − Alt_ref)/100·k_A/100. "
          "Coefficients are NOT supplied - obtain from the applicable standard/manufacturer. Kharm from IEEE C57.110 assessment or manufacturer. "
          "Where a required coefficient is missing the factor is 1 ONLY for arithmetic continuity: DerateOK = 0, the rating status cannot become 'SELECTED - adequate' and the loading check cannot PASS."),
    ("T", "Required (own) = Design·(1+margin)·(1+future)/K. Contingency: N+1 → group (Design·margins) ÷ (n−1) ÷ K; 2N → ÷ (n/2) ÷ K. "
          "Calculated minimum = max(own, contingency). Recommended standard = smallest value in active rating series ≥ minimum. This is NOT an approved selection."),
    ("T", "Loading % = Design demand ÷ evaluated rating × 100 (evaluated = Selected_kVA if entered, else recommended). Loading of derated capacity and spare kVA also shown. "
          "Utilisation per scenario U_s = S_s ÷ rating × 100."),
    ("T", "Full-load currents: I_HV = S/(√3·kV) [A, S in kVA]; I_LV = S·1000/(√3·V_LL)."),
    ("T", "Regulation: n = Design/rating, uR = uk/√(1+(X/R)²), uX = uR·X/R; ε% = n(uR·cosφ + uX·sinφ) + n²(uX·cosφ − uR·sinφ)²/200, with cosφ = aggregate demand PF."),
    ("T", "Short-circuit (preliminary): S_sc,TX = S/(uk/100); bus S_sc = series combination with MV source level (blank = infinite); Isc = c·S_sc/(√3·V). "
          "Motor-start dip ≈ S_start/(S_start + S_sc,bus)·100 (infinite-bus-behind-impedance approximation, pre-start load neglected)."),
    ("H", "6. LV distribution (sheet 7)"),
    ("T", "Feeder design kVA from the selected scenario/envelope of rows assigned to the feeder; Ib from S and V. Checks: Ib·mult ≤ In; In ≤ Iz; I2 ≤ K·Iz; "
          "ΔV% = k·Ib·L·(R cosφ + X sinφ)/V·100 (k = √3 3PH, 2 1PH); Icu ≥ Isc at TX LV terminals (conservative). Iz, R, X and device data are user inputs."),
    ("T", "Phase balance: 3PH loads split equally; 1PH L-N to assigned phase; 1PH L-L split 50/50 to the two phases (approximation of current distribution). "
          "Imbalance = (max − average)/average × 100."),
    ("H", "7. Definitions"),
    ("T", "Connected load: sum of rated electrical inputs of all installed units. Maximum (coincident) demand: highest scenario apparent power after demand factors, "
          "scenario participation and exclusivity. Demand factor (DF): ratio of absorbed input at maximum load to rated input. Duty-cycle factor: average on-time "
          "fraction during operating hours (energy only). Scenario factor: participation (simultaneity) of a load class in a scenario."),
    ("H", "8. Constants used in formulas"),
    ("T", "SQRT3 = √3 (workbook Name). 1000 = W↔kW, V·A↔kVA, kVA↔MVA and m↔km conversions. 100 = per-unit ↔ percent. 200 = constant of the percent-form regulation "
          "approximation. 2 = go-and-return conductor length for single-phase voltage drop. No other numeric constants are embedded; all criteria are Design Basis inputs."),
    ("H", "9. Limitations"),
    ("T", "• PF sign is per row (Lag/Lead); explicit kvar ratings and PF-correction banks must be represented through PF/sign. • Unbalanced-load effects on transformer capacity are not modelled "
          "beyond the imbalance check. • Ambient/altitude derating is linear with user coefficients. • Contingency assumes equal load sharing. • Short-circuit and dip are "
          "simplified, not IEC 60909 / IEEE full studies; no motor contribution. • Thermal (loading-guide) overload capability is not calculated. • Harmonic derating is a user input. "
          "• Energy is a normal-day estimate, not an 8760-h simulation. • Scenario profiles are generic assumptions requiring project review."),
    ("H", "10. Excel compatibility"),
    ("T", "Functions used: SUMIFS, COUNTIFS, IFERROR, SUMPRODUCT, INDEX, MATCH, CHOOSE, SMALL, RANK, SQRT, MAX, MIN - Excel 2010 or later (Windows/Mac). "
          "No dynamic-array / Excel-365-only functions, no VBA, no external links. Structured table references are used throughout. Recalculation is automatic."),
    ("H", "11. Extending the model"),
    ("T", "Add categories in tblCat, library items in tblLib, profiles in tblProf, transformers in tblTX and feeders in tblFdr by typing in the first row below each table. "
          "Summary/Dashboard slot lists show the first 30 categories and first 20 transformers (sheet 5) / 12 (dashboard); calculations themselves are unlimited."),
]


def build_method(wb, ws):
    title(ws, "Sheet 10 - Calculation Methodology & User Guide", f"Workbook version {VERSION}")
    ws.column_dimensions["A"].width = 160
    r = 4
    for kind, txt in METHOD:
        c = ws.cell(row=r, column=1, value=txt)
        if kind == "H":
            c.font = FN_B
            c.fill = F_SUB
        else:
            c.alignment = WRAP
            ws.row_dimensions[r].height = max(15, 15 * (len(txt) // 150 + 1))
        r += 1


def build_static_sheet(ws, title_text, sub, header, rows, widths):
    title(ws, title_text, sub)
    for j, (h, w) in enumerate(zip(header, widths)):
        c = ws.cell(row=4, column=1 + j, value=h)
        c.fill = F_HDR
        c.font = FN_HDR
        ws.column_dimensions[CL(1 + j)].width = w
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            c = ws.cell(row=5 + i, column=1 + j, value=v)
            c.alignment = WRAP
            c.border = BRD
    status_cf(ws, f"{CL(len(header) - 1)}5:{CL(len(header))}{5 + len(rows)}")


# ======================================================================= MAIN
def build(path, cfg=None, test_log=None, assumptions=None):
    cfg = cfg or {}
    wb = Workbook()
    names = ["Design Basis", "Equipment Library", "Load Schedule", "Scenario Analysis", "Summary",
             "Transformer Sizing", "LV Distribution", "Validation", "Dashboard", "Methodology", "Assumptions", "Test Log", "Lists"]
    wss = {}
    for i, n in enumerate(names):
        wss[n] = wb.active if i == 0 else wb.create_sheet()
        wss[n].title = n
    add_name(wb, "SQRT3", "SQRT(3)")
    build_lists(wb, wss["Lists"], cfg)
    build_design_basis(wb, wss["Design Basis"], cfg)
    build_library(wb, wss["Equipment Library"], cfg)
    build_scenarios(wb, wss["Scenario Analysis"], cfg)
    build_load(wb, wss["Load Schedule"], cfg)
    build_tx(wb, wss["Transformer Sizing"], cfg)
    build_fdr(wb, wss["LV Distribution"], cfg)
    pos = build_summary(wb, wss["Summary"], cfg)
    build_validation(wb, wss["Validation"], cfg)
    build_dashboard(wb, wss["Dashboard"], cfg)
    add_charts(wb, wss["Dashboard"], wss["Summary"], pos)
    build_method(wb, wss["Methodology"])
    build_static_sheet(wss["Assumptions"], "Assumptions, unresolved items & limitations",
                       "Register of model assumptions. Status to be closed by the responsible engineer.",
                       ["Ref", "Item", "Type", "Treatment in workbook", "Action required", "Status"],
                       assumptions or [], [6, 40, 16, 60, 50, 14])
    build_static_sheet(wss["Test Log"], "Independent verification test log",
                       "Executed by recalculating test-case workbooks in a spreadsheet engine and comparing to an independent Python implementation.",
                       ["Test", "Dataset", "Case", "Quantity checked", "Expected (independent)", "Workbook result", "Abs. diff", "Result"],
                       test_log or [], [8, 22, 30, 48, 20, 20, 12, 10])
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)
    return path
