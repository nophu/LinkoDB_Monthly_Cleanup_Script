# Per-report field config — taken directly from the monthly tracker.
# Maps each filename to the report name and the fields that should
# be validated for that specific file.
DATA_DIR = "xlsx"  # folder where all Excel files live
REPORT_CONFIG = {
    "Master List with Additional Fields.xlsx": {
        "name":   "Master List with Additional Fields",
        "fields": ["MapCategory"],
    },
    "r_M_FG_sys_Iu_Permit_List_Extractors.xlsx": {
        "name":   "Permit List – Extractors",
        "fields": ["Extractor ID", "Extractor Type", "Trap Size and Units", "Cleaning Frequency"],
    },
    "2025-08-05 FSE Inspections last 5 years.xlsx": {
        "name":   "FSE Inspections – Last 5 Years",
        "fields": ["ReceivingPlant", "ClassCode", "SecondClass", "TrunkLine"],
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