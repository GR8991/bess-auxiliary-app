"""JSON Schema (Draft 2020-12) of the Power System BOP Engineer AI-Agent handoff package."""
HANDOFF_SCHEMA_VERSION = "1.0"

_tagged = {"type": "object", "required": ["value", "source"],
           "properties": {"value": {}, "source": {"enum": ["user_input", "workbook_default", "derived", "calculated", "missing"]},
                          "unit": {"type": ["string", "null"]}, "check": {"type": ["string", "null"]}}}
_rows = {"type": "array", "items": {"type": "object"}}

HANDOFF_SCHEMA = {
    "$schema": "https://json-schema.org/draft/2020-12/schema",
    "$id": "bess-aux-bop-handoff-1.0",
    "title": "BESS Auxiliary Engineering - BOP Engineer AI Agent handoff package",
    "description": "Deterministic results of the BESS auxiliary calculation engine (workbook v1.1 methodology). Every value carries a "
                   "source tag. The receiving agent must not treat calculated values as approved design, and must not fill missing inputs.",
    "type": "object",
    "required": ["package_type", "handoff_schema_version", "identification", "project_information", "design_basis", "equipment_library",
                 "load_schedule", "operating_scenarios", "transformer_sizing", "lv_distribution", "assumptions", "validation",
                 "outstanding_data_requirements", "engineering_limitations", "engineering_recommendations", "methodology",
                 "input_completion", "overall_disposition", "value_source_legend"],
    "properties": {
        "package_type": {"const": "BESS_AUX_BOP_HANDOFF"},
        "handoff_schema_version": {"type": "string"},
        "identification": {"type": "object", "required": ["project_id", "revision", "generated_utc", "application_version",
                                                         "calculation_engine_version", "source_workbook_version", "project_schema_version",
                                                         "handoff_schema_version"]},
        "project_information": {"type": "object", "additionalProperties": _tagged},
        "design_basis": {"type": "object", "additionalProperties": _tagged},
        "equipment_library": {"type": "object", "required": ["source", "items"], "properties": {"items": _rows}},
        "load_schedule": {"type": "object", "required": ["user_inputs", "calculated", "input_formula_trace"],
                          "properties": {"user_inputs": _rows, "calculated": _rows, "input_formula_trace": _rows}},
        "operating_scenarios": {"type": "object", "required": ["definitions", "profiles", "results"]},
        "transformer_sizing": {"type": "object", "required": ["user_inputs", "calculated", "centralized_comparison"]},
        "lv_distribution": {"type": "object", "required": ["user_inputs", "calculated", "application_feeder_phase_balance"]},
        "assumptions": _rows,
        "validation": {"type": "object", "required": ["register", "findings", "status_counts"],
                       "properties": {"register": _rows, "findings": _rows}},
        "outstanding_data_requirements": _rows,
        "engineering_limitations": {"type": "array", "items": {"type": "string"}},
        "engineering_recommendations": _rows,
        "methodology": {"type": "object"},
        "input_completion": _rows,
        "overall_disposition": {"type": "object", "required": ["status", "note"],
                                "properties": {"status": {"enum": ["INPUTS INCOMPLETE", "CALCULATION COMPLETED — REVIEW REQUIRED",
                                                                   "PRELIMINARY ENGINEERING REVIEW", "READY FOR INDEPENDENT ENGINEERING REVIEW"]}}},
        "value_source_legend": {"type": "object"},
    },
}
