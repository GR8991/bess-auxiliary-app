# User guide — BESS Auxiliary Engineering

## Signing in
Enter your approved company email address. This is a **basic allowlist check, not identity verification**, and no password is used.
Sessions expire after the configured period of inactivity. **Log out** clears everything in your session.

## Your data
The application does not store projects. Use **Save project (JSON)** in the sidebar often, and before closing the browser. Reopen
a project with *Project Overview → Upload existing project*. When a file comes from an older application or engine version, you
are warned and the results are recalculated with the current engine.

## Workflow

| Step | Page | What to do |
|---|---|---|
| 1 | Project Overview | Create a project, upload a JSON project, or import a v1.1 Excel workbook (input cells only). Enter project information. 🔗 fields are the Design Basis inputs themselves. |
| 2 | Design Basis | Complete every ★ input. Enter margins, limits and derating coefficients from the applicable standard or manufacturer. Percent inputs are entered in %. Each field shows its workbook input check. |
| 3 | Equipment Library | Enter datasheet ratings with a source reference. Ratings are blank by default and none are invented. Manufacturer and DeratingNote are information only. |
| 4 | Auxiliary Load Schedule | One row per equipment type and supply point. Select a LibID to inherit library data. Quantities accept formulas such as `=DB_N_Encl/2`. Choose Lag or Lead PF. Give mutually exclusive loads the same ExclGroup (same TX and feeder). Press **Apply changes**. Deleting rows needs a confirmation tick. |
| 5 | Scenario Analysis | Review the participation profiles, confirm them in the Design Basis, and compare scenarios. |
| 6 | Transformer Sizing | Define transformers, redundancy groups and schemes, impedance and X/R. Compare the **calculated minimum**, the **recommended standard** (not approved) and your **selected rating**. |
| 7 | LV Distribution | Enter feeder data (Iz, R, X, In, I2/In, Icu) from standard tables and manufacturer data. These are screening checks only. |
| 8 | Engineering Validation | Filter the V01–V28 register and the row findings by status, severity and category. |
| 9 | Engineering Summary | Sections A–G and the overall disposition. |
| 10 | AI Agent Handoff | Download the schema-validated JSON package for the BOP Engineer AI Agent. |
| 11 | Reports and Exports | Project JSON, handoff JSON, PDF, Excel results, v1.1 formula workbook, CSV tables. |
| 12 | Application Settings | Days per year, breakdown scenario, centralized-comparison type, user rating series. |

## Reading the statuses

| Status | Meaning |
|---|---|
| PASS | The workbook check is satisfied with the data entered. This is **not** approval or compliance. |
| FAIL | The check is violated and must be resolved. |
| REVIEW | An engineering judgement or review is needed. |
| INCOMPLETE | Required input is missing. |
| NOT VERIFIED | A standard, code or AHJ item needs verification against licensed documents. |
| NOT APPLICABLE | The check does not apply. V25 always shows this in the app because the app has no display-slot limits. |

Each register item also shows the original **workbook result** for traceability. The disposition sequence is:
INPUTS INCOMPLETE → CALCULATION COMPLETED — REVIEW REQUIRED → PRELIMINARY ENGINEERING REVIEW → READY FOR INDEPENDENT ENGINEERING REVIEW.
No higher status exists.

## Key methodology points (workbook v1.1, reproduced exactly)
- **Power aggregation:** S = √(ΣP² + ΣQ²) with signed Q. Leading loads reduce the kvar total (F-06).
- **Exclusivity:** within an ExclGroup, the member with the highest kVA in each scenario counts, with its own P and Q (F-01). Members must share one TX and feeder (F-02).
- **Derating:** missing derating data blocks "SELECTED – adequate" and a PASS on loading and selection (F-04).
- **Screening-level results:** starting/inrush, regulation, fault level and voltage drop are screening only (F-05). Regulation and voltage drop use the aggregate running PF (R-1). The transient is site-wide (R-2).
- **Empty feeder table:** as in the workbook, it is evaluated as one blank row, so V19–V22 show REVIEW (finding E-1).
- **Application-level additions:** findings listed per row, feeder phase balance, display statuses and disposition are marked **NOT IN v1.1**.

## Detailed studies always required for final design
Short-circuit (IEC 60909 / IEEE), load flow, protection coordination and arc flash, motor starting and restoration, voltage-drop verification
against the final cable schedule, harmonic assessment (IEEE C57.110 / IEEE 519), and verification of standards and AHJ requirements.
