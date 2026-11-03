"""Project data model (JSON-serialisable dict) and validation.

A project holds INPUTS ONLY plus metadata; every calculated value is re-derived by the engine.
Unknown keys are preserved (never silently discarded).
"""
from __future__ import annotations

import copy
import datetime
import uuid

from calculations import reference_data as R
from calculations.constants import ENGINE_VERSION, SOURCE_WORKBOOK_VERSION
from calculations.design_basis import default_design_basis

APP_VERSION = "1.0.0"
SCHEMA_VERSION = "1.0"
SUPPORTED_SCHEMA_MAJOR = "1"

# A-3: informational Project Overview fields (no calculation effect)
PROJECT_INFO_FIELDS = [
    ("client", "Client"), ("country_grid_region", "Country / grid region"), ("poi_voltage_kv", "POI voltage (kV)"),
    ("reviewed_by", "Reviewed by"), ("date", "Date"),
]
# Overview fields that ARE Design-Basis inputs (single source of truth - no duplicate inputs)
OVERVIEW_DB_LINKS = [("Project", "Project name"), ("Location", "Project location"), ("BESS_MW", "Project capacity (MW)"),
                     ("BESS_MWh", "BESS energy capacity (MWh)"), ("VLL", "Auxiliary system voltage, L-L (V) = Design Basis LV voltage"),
                     ("Rev", "Project revision"), ("Author", "Prepared by")]

LIBRARY_EXTRA_COLUMNS = ["Manufacturer", "DeratingNote"]   # A-4 (informational; DeratingNote has no calculation effect)
TABLE_INPUTS = {
    "loads": R.TABLE_COLUMNS["tblLoad"]["inputs"],
    "transformers": R.TABLE_COLUMNS["tblTX"]["inputs"],
    "feeders": R.TABLE_COLUMNS["tblFdr"]["inputs"],
    "library": R.TABLE_COLUMNS["tblLib"]["inputs"] + LIBRARY_EXTRA_COLUMNS,
    "categories": R.TABLE_COLUMNS["tblCat"]["inputs"],
    "profiles": R.TABLE_COLUMNS["tblProf"]["inputs"],
    "scenarios": R.TABLE_COLUMNS["tblScen"]["inputs"],
}
NUMERIC = {
    "loads": {"QtyInst", "QtyDuty", "VOverride", "Rating", "PF", "Eff", "DF", "DutyF", "Hours", "StartMult", "StartPF", "THDi"},
    "transformers": {"HV_kV", "LV_VLL", "LV_VLN", "Z_pct", "XR", "Kharm_in", "InrushMult", "Selected_kVA"},
    "feeders": {"Length_m", "Iz_A", "R_ohm_km", "X_ohm_km", "In_A", "I2_ratio", "Icu_kA"},
    "library": {"RatedValue", "PF", "Eff", "StartMult", "StartPF"},
    "profiles": set(R.LISTS["lstScenCodes"]),
}
EXPRESSION_FIELDS = {"loads": {"QtyInst", "QtyDuty", "Rating"}}   # A-6: Excel-style '=...' allowed (UI); engine accepts any
DB_TEXT_KEYS = {"Project", "Location", "DocNo", "Rev", "Author", "Checker", "Tech", "Jurisdiction", "Earthing", "Arrangement",
                "StdSeries", "DesignScen", "ProfRev", "StdRev"}
DB_LIST = {"Jurisdiction": "lstJuris", "Freq": "lstFreq", "Earthing": "lstEarth", "Arrangement": "lstArr", "StdSeries": "lstSeries",
           "DesignScen": "lstDesignScen", "ProfRev": "lstYN", "StdRev": "lstYN"}
DB_MANDATORY = {"Jurisdiction", "Freq", "MV", "VLL", "Arrangement", "Tamb", "Alt"}
DB_PERCENT = {"Margin", "Future", "MaxLoad", "MaxLoadCont", "RegLim", "DipLim", "VDLim", "NLThr", "ImbLim"}
TOP_KEYS = ["schema_version", "metadata", "project_info", "design_basis", "categories", "library", "scenarios", "profiles", "loads",
            "transformers", "feeders", "rating_series", "settings"]
DEFAULT_SETTINGS = dict(days_per_year=365, breakdown_scenario="MAXSIM", central_type="Liquid-filled")


def now_iso():
    return datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds")


def new_project(name: str = "", created_by: str = "") -> dict:
    db = default_design_basis()
    db["Project"] = name or None
    lib = [dict(row, Manufacturer=None, DeratingNote=None) for row in copy.deepcopy(R.LIBRARY)]
    return dict(
        schema_version=SCHEMA_VERSION,
        metadata=dict(project_id=str(uuid.uuid4()), revision="Rev00", created=now_iso(), updated=now_iso(), created_by=created_by,
                      app_version=APP_VERSION, engine_version=ENGINE_VERSION, source_workbook_version=SOURCE_WORKBOOK_VERSION,
                      status="Draft"),
        project_info={k: None for k, _ in PROJECT_INFO_FIELDS},
        design_basis=db,
        categories=copy.deepcopy(R.CATEGORIES), library=lib, scenarios=copy.deepcopy(R.SCENARIOS), profiles=copy.deepcopy(R.PROFILES),
        loads=[], transformers=[], feeders=[], rating_series=copy.deepcopy(R.RATING_SERIES), settings=dict(DEFAULT_SETTINGS),
    )


def _is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def validate_project(p) -> tuple[list, list]:
    """Return (errors, warnings). Errors block loading; warnings are shown to the user."""
    errors, warnings = [], []
    if not isinstance(p, dict):
        return ["Project file is not a JSON object"], []
    sv = str(p.get("schema_version", ""))
    if not sv:
        errors.append("schema_version missing")
    elif sv.split(".")[0] != SUPPORTED_SCHEMA_MAJOR:
        errors.append(f"Unsupported schema version {sv} (this application reads {SUPPORTED_SCHEMA_MAJOR}.x)")
    elif sv != SCHEMA_VERSION:
        warnings.append(f"Project schema {sv} differs from application schema {SCHEMA_VERSION}")
    for k in TOP_KEYS:
        if k not in p:
            errors.append(f"Required section '{k}' missing")
    unknown = [k for k in p if k not in TOP_KEYS and k != "results_snapshot"]
    if unknown:
        warnings.append(f"Unknown sections preserved unchanged: {unknown}")
    if errors:
        return errors, warnings
    if not isinstance(p["design_basis"], dict):
        errors.append("design_basis must be an object")
    for key in TABLE_INPUTS:
        rows = p.get(key)
        if not isinstance(rows, list) or not all(isinstance(r, dict) for r in rows):
            errors.append(f"'{key}' must be a list of objects")
            continue
        allowed = set(TABLE_INPUTS[key]) | {"_extra", "_padding"}
        for i, r in enumerate(rows):
            extra = [c for c in r if c not in allowed]
            if extra:
                warnings.append(f"{key}[{i + 1}]: unknown fields preserved {extra}")
            for c in NUMERIC.get(key, ()):
                v = r.get(c)
                if v is None or v == "" or _is_num(v):
                    continue
                if isinstance(v, str) and v.startswith("="):
                    continue
                warnings.append(f"{key}[{i + 1}].{c}: non-numeric value {v!r} (the engine treats it as Excel would)")
    meta = p.get("metadata") or {}
    for k in ("project_id", "revision"):
        if not meta.get(k):
            errors.append(f"metadata.{k} missing")
    if meta.get("engine_version") and meta["engine_version"] != ENGINE_VERSION:
        warnings.append(f"Project saved with engine {meta['engine_version']}; results recalculated with engine {ENGINE_VERSION}")
    if meta.get("app_version") and meta["app_version"] != APP_VERSION:
        warnings.append(f"Project saved with application {meta['app_version']}; current application {APP_VERSION}")
    if meta.get("source_workbook_version") and meta["source_workbook_version"] != SOURCE_WORKBOOK_VERSION:
        warnings.append(f"Project references workbook {meta['source_workbook_version']}; engine reproduces v{SOURCE_WORKBOOK_VERSION}")
    return errors, warnings


def engine_input(p: dict) -> dict:
    """The subset the calculation engine consumes (informational fields are excluded)."""
    return {k: p.get(k) for k in ("design_basis", "categories", "library", "scenarios", "profiles", "loads", "transformers", "feeders",
                                  "rating_series", "settings")}
