"""Stateless project storage: JSON download / upload only (no server-side persistence)."""
from __future__ import annotations

import json
import re

from calculations.constants import ENGINE_VERSION, SOURCE_WORKBOOK_VERSION
from data_models.project_schema import APP_VERSION, SCHEMA_VERSION, validate_project, now_iso


def filename(project: dict) -> str:
    name = (project.get("design_basis") or {}).get("Project") or "Project"
    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", str(name)).strip("_") or "Project"
    rev = re.sub(r"[^A-Za-z0-9]+", "", str((project.get("metadata") or {}).get("revision") or "Rev00")) or "Rev00"
    return f"{safe}_BESS_Auxiliary_Design_{rev}.json"


def to_json(project: dict, results: dict | None = None) -> bytes:
    """Serialise inputs + metadata. A small results snapshot is included for reference only;
    it is NEVER read back - results are always recalculated on load."""
    p = json.loads(json.dumps(project, default=str))
    meta = p.setdefault("metadata", {})
    meta.update(updated=now_iso(), app_version=APP_VERSION, engine_version=ENGINE_VERSION,
                source_workbook_version=SOURCE_WORKBOOK_VERSION)
    p["schema_version"] = SCHEMA_VERSION
    if results:
        app = results.get("application", {})
        k = results["workbook"]["summary"]["kpi"]
        p["results_snapshot"] = dict(note="Informational snapshot; recalculated on load", disposition=app.get("disposition"),
                                     peak_kVA=k.get("SUM_PeakkVA"), peak_scenario=k.get("SUM_PeakName"),
                                     status_counts=app.get("status_counts"))
    return json.dumps(p, indent=2, ensure_ascii=False, default=str).encode("utf-8")


def from_json(data: bytes | str) -> tuple[dict | None, list, list]:
    """Return (project, errors, warnings). Unknown content is preserved, never discarded."""
    try:
        p = json.loads(data.decode("utf-8") if isinstance(data, (bytes, bytearray)) else data)
    except (UnicodeDecodeError, json.JSONDecodeError) as e:
        return None, [f"File is not valid JSON: {e}"], []
    errors, warnings = validate_project(p)
    if errors:
        return None, errors, warnings
    if "results_snapshot" in p:
        warnings.append("Saved results snapshot ignored - all outputs recalculated with the current engine")
    return p, errors, warnings
