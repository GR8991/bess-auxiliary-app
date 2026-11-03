# Project status and phase completion reports — BESS Auxiliary Engineering App v1.0.0

Status date: 2026-11-03. All test numbers below come from executed runs (`python -m pytest` → **54 passed**).

## Phase 1 — Workbook reverse engineering ✅
- **Completed:** architecture report, formula inventory (473 items), named ranges (124), data validation (75 rules), dependency diagram, validation mapping, implementation plan.
- **Files:** `docs/PHASE1_REVERSE_ENGINEERING_REPORT.md`, `docs/formula_inventory_v1.1.csv`, `docs/named_ranges_v1.1.csv`, `docs/data_validation_v1.1.csv`.
- **Decisions received:**
  - A-1, A-2, A-3, A-4, A-7 and A-9 approved.
  - A-5 approved as extra screening.
  - A-6 changed: reproduce input-cell formulas.
  - A-8 approved with conditions.
  - R-1 to R-6 to be reproduced exactly.

## Phase 2 — Calculation engine and equivalence ✅
- **Completed:**
  - `calculations/` engine: all 58 tblLoad, 69 tblTX and 31 tblFdr calculated columns, scenarios, Summary, centralized comparison, V01–V28 and standards applicability.
  - Excel semantics layer and the A-6 input-formula evaluator.
  - A-7 Excel importer and the reference-baseline builder.
- **Equivalence:**
  - **119,876 comparisons, 0 failures** across 14 datasets (12 v1.1 datasets plus DS11 with 31 categories and DS12 with phase-balance edge cases), the delivered template and the delivered worked example.
  - Two input paths were used: dataset JSON and Excel import.
  - Formula identity with the delivered template was verified for all 30 reference workbooks.
  - **All 382 calculated items are mapped and have passing comparisons** (`docs/formula_mapping_v1.1.csv`).
- **Unit tests (hand calculations, independent of the workbook):** Excel semantics, input formulas, load equations, scenarios, transformer sizing, N+1/2N, regulation/Isc/dip, phase balance, F-01…F-06, more than 20 TX, more than 30 categories, A-5, A-9, disposition.
- **Reports:** `docs/CALCULATION_EQUIVALENCE_REPORT.md`, `docs/equivalence_results.csv`.

## Phase 3 — Framework, allowlist gate, JSON storage ✅
- **Completed:**
  - `app.py` with `st.navigation`, a sidebar (project, user, disposition, Save JSON, unsaved-changes warning, logout), and a page guard on every page.
  - `authentication/access_control.py`: normalisation, generic denial, 5-attempt lockout, inactivity timeout, logout that clears the session. Addresses come from `st.secrets` only.
  - `services/project_storage.py`: schema validation, version warnings, preservation of unknown data, recalculation on load.
- **Tests:** gate, lockout, timeout, logout, missing secrets, JSON round-trip, rejection of bad files.

## Phase 4 — Engineering modules ✅
- **Completed:** Project Overview (A-3 fields; Aux voltage linked to DB_VLL), Design Basis (sections, ★ mandatory, % inputs, per-field checks), Equipment Library (+Manufacturer, DeratingNote, derived kW/kVA), Load Schedule, Scenario Analysis, Transformer Sizing (three separate ratings, DerateOK), LV Distribution (screening banner, A-5 feeder balance).
- **Table editing:** explicit Apply, a confirmation tick for deleting rows, row duplication, lossless `_extra` fields.

## Phase 5 — Summary, validation dashboard, AI handoff ✅
- **Completed:**
  - Engineering Summary sections A–G (completion uses four statuses, not a percentage).
  - Validation dashboard with filters, showing both the workbook result and the application status.
  - Handoff JSON with a JSON Schema (Draft 2020-12) and per-value source tags.

## Phase 6 — Visual refinement ✅ (code) / ⚠️ not viewed in a real browser
- **Completed:** dark navy design system (CSS), branded login page with a generic BESS illustration, Plotly charts 1–8 built only from engine outputs.
- **Checked:** the PDF report was rendered and visually reviewed.

## Phase 7 — QA, exports, documentation ✅
- **Exports:** project JSON, handoff JSON, Excel results, v1.1 formula workbook (round-trip test: export → LibreOffice → equal to the engine), PDF, CSV.
- **Docs:** README, deployment guide, user guide.

## Open items (honest list)

| ID | Item | Owner |
|---|---|---|
| D-1 | Microsoft Excel desktop recalculation not verified; the baseline is LibreOffice Calc. | User: open the exported v1.1 workbook in Excel and recalculate. |
| UI-1 | **Streamlit and Plotly could not be installed in the build environment (PyPI blocked).** The UI was executed only against a stand-in (`tests/ui_stubs`) and has not yet been run in a real browser. | First real run (`streamlit run app.py`), then deployment checks in `DEPLOYMENT_GUIDE.md` §6. |
| E-1 | v1.1 behaviour: blank feeder rows / an empty feeder table make V19–V22 REVIEW, and V20 lacks a "no feeders" guard. Reproduced, not changed. | User decision. |
| F-07 | Standards editions and clauses need verification against licensed copies. | Responsible engineer. |
| SEC-1 | The email allowlist gate is not identity verification. | Consider a private Cloud app or OIDC in a later release. |
| REV-1 | Independent third-party engineering review of the methodology. | User. |
