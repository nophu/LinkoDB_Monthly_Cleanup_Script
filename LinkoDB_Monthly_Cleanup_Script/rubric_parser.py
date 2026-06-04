import re
import json
import pandas as pd

# green tabs only
NEEDED_SHEETS = [
    'Key',
    'Tables Extractor IDs, Type',
    '6Other Codes_MapCategory_Events',
    'FS - Trap Size & Units',
    'FS - Cleaning Frequency',
]

# 6Other Codes sheet has side-by-side columns where each "New ___" column
OTHER_CODES_COLUMNS = {
    "new class codes":    "ClassCode",
    "new secondclass":    "SecondClass",
    "new mapcategory":    "MapCategory",
    "new eventtypeabbrv": "EventTypeAbbrv",
}

# fields owned by the 6Other Codes sheet — the generic reader must NOT touch these
OTHER_CODES_FIELDS = set(OTHER_CODES_COLUMNS.values()) | {"ReceivingPlant", "TrunkLine"}

# fields owned by the Tables sheet — generic reader must not touch these either
TABLES_FIELDS = {"Extractor ID", "Extractor Type"}

def parse_rubric(filepath):
    print(f"\nReading rubric: {filepath}")
    fp = filepath.lower()
    if fp.endswith(".xlsx"):    engine = 'openpyxl'
    elif fp.endswith(".xls"):    engine = 'xlrd'
    else:   raise ValueError(f"Unrecognized file type: {filepath}")
    wb = pd.ExcelFile(filepath, engine=engine)
    sheet_names = wb.sheet_names
    rubric = {"column_mapping": {}, "valid_values": {}, "value_patterns": {}}

    # learn column-name mapping from the Key sheet
    key_sheet = _find_sheet(wb, "Key")
    if key_sheet is not None:
        for _, row in key_sheet.iterrows():
            values = [str(v).strip() for v in row if str(v).strip() not in ("", "nan")]
            if len(values) >= 2:
                messy, correct = values[0], values[1]
                if messy.lower() in ("download data header", "header", "field"):   continue
                rubric["column_mapping"][messy] = correct
        print(f"Column mappings learned: {len(rubric['column_mapping'])}")
    else:  print("WARNING: Could not find 'Key' sheet")

    # extract valid values from each green rule sheet
    rule_sheets = [s for s in sheet_names if s in NEEDED_SHEETS]
    print(f"Rule sheets used: {rule_sheets}")

    for sheet_name in rule_sheets:

        # this sheet needs its own dedicated reader
        if "Other Codes" in sheet_name:   _extract_other_codes(wb, sheet_name, rubric)

            # only read the NEW Linko pick list (ignore the Current/old scheme)
        elif "Tables Extractor" in sheet_name: _extract_tables(wb, sheet_name, rubric)

            # all other green sheets use the generic single-column reader
        else: _extract_valid_values(wb, sheet_name, rubric)
    _build_patterns(rubric)

    # trap Size and Units uses a conditional size-based rule, not a static list.
    rubric["valid_values"]["Trap Size and Units"] = []
    with open("output/rubric.json", "w") as f:    json.dump(rubric, f, indent=2)
    print("Saved: output/rubric.json")
    return rubric

# dedicated reader for the "Tables Extractor IDs, Type" sheet
def _extract_tables(wb, sheet_name, rubric):
    df = wb.parse(sheet_name, dtype=str, header=None)

    # find the boundary row where the OLD (Current) section starts
    boundary = len(df)
    for r in range(len(df)):
        rowtext = " ".join(str(c).strip().lower() for c in df.iloc[r])
        if "current linko pick" in rowtext:
            boundary = r
            break

    # extractor IDs live in column 1 — keep only cells that look like "EX###"
    ext_ids = []
    for r in range(boundary):
        cell = str(df.iloc[r, 1]).strip()
        if re.match(r"^EX\s*\d", cell, re.IGNORECASE):   ext_ids.append(cell)
    if ext_ids:  rubric["valid_values"]["Extractor ID"] = sorted(set(ext_ids))

    # extractor Types live in column 7 of the NEW section — only ABOVE-FLOOR / IN-FLOOR / OUTDOOR
    ext_types = []
    if df.shape[1] > 7:
        for r in range(boundary):
            cell = str(df.iloc[r, 7]).strip()
            if cell in ("", "nan"):  continue
            cl = cell.lower()
            if "extractor type" in cl or "description" in cl or "location" in cl:  continue
            if len(cell) > 30:    continue
            ext_types.append(cell)
    if ext_types:  rubric["valid_values"]["Extractor Type"] = sorted(set(ext_types))

# dedicated reader for the 6Other Codes sheet
def _extract_other_codes(wb, sheet_name, rubric):
    df = wb.parse(sheet_name, dtype=str, header=None)

    # find the row that contains the "New ___" column headers
    header_row = 0
    best = 0
    for r in range(min(5, len(df))):
        hits = sum(1 for c in df.iloc[r] if str(c).strip().lower() in OTHER_CODES_COLUMNS)
        if hits > best:
            best = hits
            header_row = r

    # read each "New ___" column — values below the header are the valid list
    for col_idx in range(df.shape[1]):
        label = str(df.iloc[header_row, col_idx]).strip().lower()
        if label not in OTHER_CODES_COLUMNS:   continue
        field = OTHER_CODES_COLUMNS[label]
        vals = []
        for r in range(header_row + 1, len(df)):
            cell = str(df.iloc[r, col_idx]).strip()
            if cell in ("", "nan"):  continue
            if len(cell) > 40:  continue
            if cell.lower().startswith("delete entry"):   continue
            vals.append(cell)
        if vals:   rubric["valid_values"][field] = sorted(set(vals))

    # trunkLine rule: "delete entry and leave field BLANK"
    rubric["valid_values"]["TrunkLine"] = []

    # scan for the "new receiving plant" sub-header, then collect values below it.
    receiving_vals = []
    for col_idx in range(df.shape[1]):
        for row_idx in range(len(df)):
            cell = str(df.iloc[row_idx, col_idx]).strip().lower()
            if cell == "new receiving plant":
                for r in range(row_idx + 1, min(row_idx + 10, len(df))):
                    val = str(df.iloc[r, col_idx]).strip()
                    if val in ("", "nan"):  continue
                    if val.lower().startswith("delete entry"):   continue
                    if len(val) > 60: continue
                    receiving_vals.append(val)
                break
    if receiving_vals:   rubric["valid_values"]["ReceivingPlant"] = sorted(set(receiving_vals))

# generic single-column reader for Tables / FS - Trap Size / FS - Cleaning Frequency
def _extract_valid_values(wb, sheet_name, rubric):
    try: df = wb.parse(sheet_name, dtype=str, header=None)
    except Exception as e:
        print(f"   WARNING: Could not read sheet '{sheet_name}': {e}")
        return

    known_fields = list(rubric["column_mapping"].values())
    for row_idx, row in df.iterrows():
        for col_idx, cell in enumerate(row):
            cell_str = str(cell).strip()
            matched_field = _match_field_name(cell_str, known_fields)

            # skip fields that belong to the 6Other Codes or Tables sheets
            if matched_field and matched_field not in OTHER_CODES_FIELDS and matched_field not in TABLES_FIELDS:
                values = _collect_column_values(df, row_idx, col_idx)
                if values:
                    existing = rubric["valid_values"].get(matched_field, [])
                    rubric["valid_values"][matched_field] = sorted(set(existing + values))

def _find_sheet(wb, partial_name):
    for name in wb.sheet_names:
        if partial_name.lower() in name.lower():
            try:  return wb.parse(name, dtype=str)
            except Exception as e:  print(f"   WARNING: Could not read sheet '{name}': {e}")
    return None

def _match_field_name(cell_str, known_fields):
    cell_norm = re.sub(r"[^a-z0-9 ]", "", cell_str.lower()).strip()
    for field in known_fields:
        field_norm = re.sub(r"[^a-z0-9 ]", "", field.lower()).strip()
        if field_norm and field_norm in cell_norm:    return field
    return None

def _collect_column_values(df, start_row, col_idx):
    values = []
    field_names = [
        "extractor id", "extractor type", "trap size and units",
        "cleaning frequency", "receivingplant", "classcode",
        "secondclass", "eventtypeabbrv", "mapcategory",
    ]
    junk_words = [
        "total", "rows", "script", "number", "delete", "pick list",
        "description", "comments", "current", "future", "download",
        "header", "series", "interceptor", "separator", "trap -",
        "shared", "grease", "multiple", "facilities", "retired",
        "removed", "unknown", "effluent", "verified", "inactive",
        "extractor ids", "etc",
    ]
    for row_idx in range(start_row + 1, min(start_row + 50, len(df))):
        cell = str(df.iloc[row_idx, col_idx]).strip()
        if cell in ("", "nan"):  continue
        if len(cell) > 40: continue
        if cell.isdigit():  continue
        cl = cell.lower()
        if cl in field_names:  continue
        if any(j in cl for j in junk_words):  continue
        values.append(cell)
    return values

def _build_patterns(rubric):
    for field, values in rubric["valid_values"].items():
        if not values: continue
        p = _infer_pattern(field, values)
        if p: rubric["value_patterns"][field] = p

def _infer_pattern(field, values):
    clean = [str(v).strip() for v in values if str(v).strip()]
    if not clean:  return None
    if "extractor id" in field.lower(): return r"^EX\s*[-]?\s*\d+"
    escaped = [re.escape(v) for v in clean if len(v) <= 40]
    if escaped: return r"^(" + "|".join(escaped) + r")$"
    return None