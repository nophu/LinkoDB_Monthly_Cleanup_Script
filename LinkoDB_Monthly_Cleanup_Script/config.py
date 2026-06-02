# Per-report field config — taken directly from the monthly tracker.
# Each entry maps a filename to the report name and the fields that
# should actually be validated for that file.
DATA_DIR = "xlsx"
REPORT_CONFIG = {
    "r_M_FG_sys_Iu_Permit_List_Extractors.xlsx": {
        "name":   "Permit List – Extractors",
        "fields": ["Extractor ID", "Extractor Type", "Trap Size and Units", "Cleaning Frequency"],
    },
    "2025-08-05_FSE_Inspections_last_5_years.xlsx": {
        "name":   "FSE Inspections – Last 5 Years",
        "fields": ["ReceivingPlant", "ClassCode", "SecondClass", "TrunkLine"],
    },
    "Master_List_with_Additional_Fields.xlsx": {
        "name":   "Master List with Additional Fields",
        "fields": ["MapCategory"],
    },
    "r_M_AG_sys_Iu_Extract_Summary.xlsx": {
        "name":   "AG Extract Summary",
        "fields": ["Extractor ID", "Extractor Type"],
    },
    "r_sys_Events_Inspection_Details.xlsx": {
        "name":   "Inspection Events",
        "fields": ["EventTypeAbbrv"],
    },
}