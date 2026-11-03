"""Engine constants. Scenario codes are fixed by the v1.1 workbook (lstScenCodes); their order is
significant (CHOOSE/MATCH in Design_kVA, Gov tie order)."""
from . import reference_data as R

SC = list(R.LISTS["lstScenCodes"])          # ['NORM','MAXCOOL','MAXSIM','STBY','MAINT','EMERG','BLACK']
DESIGN_SCEN = list(R.LISTS["lstDesignScen"])
ENGINE_VERSION = "2.0.0"                     # Python calculation engine
SOURCE_WORKBOOK_VERSION = R.SOURCE_WORKBOOK_VERSION
TOL_REL = 1e-9                               # equivalence tolerance (relative, with abs floor 1e-9)
SUMMARY_CAT_SLOTS = 30                       # workbook Summary category slots (V25 / equivalence)
SUMMARY_TX_SLOTS = 20                        # workbook Summary / Scenario matrix / Dashboard TX slots
