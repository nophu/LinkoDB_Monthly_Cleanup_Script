# Each report is identified by a stable KEY. "match" lists keywords used to
# recognize a file you pick in the file explorer (case-insensitive, so the
# dated FSE file keeps working month to month). "fields" is what to check.
REPORT_CONFIG = {
    "permit_list": {
        "name":   "Permit List – Extractors",
        "match":  ["permit"],
        "fields": ["Extractor ID", "Extractor Type", "Trap Size and Units", "Cleaning Frequency"],
    },
    "fse": {
        "name":   "FSE Inspections – Last 5 Years",
        "match":  ["fse"],
        "fields": ["ReceivingPlant", "ClassCode", "SecondClass", "TrunkLine"],
    },
    "master": {
        "name":   "Master List with Additional Fields",
        "match":  ["master"],
        "fields": ["MapCategory"],
    },
    "ag": {
        "name":   "AG Extract Summary",
        "match":  ["extract_summary", "extract summary", "ag_sys"],
        "fields": ["Extractor ID", "Extractor Type"],
    },
    "events": {
        "name":   "Inspection Events",
        "match":  ["events_inspection", "inspection_details", "_events_"],
        "fields": ["EventTypeAbbrv"],
    },
}


# Given a filename, return (key, cfg) for the report it belongs to, or (None, None).
def match_report(filename):
    fn = filename.lower()
    for key, cfg in REPORT_CONFIG.items():
        if any(kw in fn for kw in cfg["match"]):  return key, cfg
    return None, None