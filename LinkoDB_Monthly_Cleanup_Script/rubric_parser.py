import re
import json
import pandas as pd

def parse_rubric(filepath):
    print(f"\nReading rubric: {filepath}")
    wb  = pd.ExcelFile(filepath)
    sheet_names = wb.sheet_names
    print(f"Sheets found: {sheet_names}")

    # column_mapping = converts messy name to correct name
    # valid_values = field name: ALL allowed values
    # value_patterns = field name for regex pattern built based off values
    rubric = { "column_mapping": {},   "valid_values":   {},   "value_patterns": {} }

    # learn column name mappings from the rubric sheet, or the key
    key_sheet = _find_sheet(wb, "Key")
    if key_sheet is not None:
        for _, row in key_sheet.iterrows():
            # grab both columns by positions in case headers shift
            values = [str(v).strip() for v in row if str(v).strip() not in ("", "nan")]

            # need at least two values, one for the messy name, and one for the correct name
            if len(values) >= 2:
                messy   = values[0]
                correct = values[1]

                # skip any row  that looks like a header row itself
                if messy.lower() in ("download data header", "header", "field"): continue
                rubric["column_mapping"][messy] = correct


        print(f"Column mappings learned: {len(rubric['column_mapping'])}")
    else: print("WARNING: Could not find 'Key' sheet")

    # learn valid values from every sheet that uses the label, "Tables" or "FS -"
    needed_sheets = ['Tables Extractor IDs, Type', 'Key', '6Other Codes_MapCategory_Events', 'FS - Trap Size & Units', 'FS - Cleaning Frequency']
    rule_sheets = [s for s in sheet_names if s in needed_sheets]
    print(f"Rule sheets found: {rule_sheets}")

    for sheet_name in rule_sheets: _extract_valid_values(wb, sheet_name, rubric)
    _build_patterns(rubric)

    # save everything to a json file
    with open("output/rubric.json", "w") as f: json.dump(rubric, f, indent=2)
    print("Saved: output/rubric.json")
    return rubric

# helper function for finding partial names in a sheet (name matching that is case-INSENSITIVE)
def _find_sheet(wb, partial_name):
    for name in wb.sheet_names:
        if partial_name.lower() in name.lower():
            try: return wb.parse(name, dtype=str)
            except Exception as e: print(f"   WARNING: Could not read sheet '{name}': {e}")
    return None

# extract valid values from a rule sheet
def _extract_valid_values(wb, sheet_name, rubric):
    # read without assuming a header, so we can get all values
    try: df = wb.parse(sheet_name, dtype=str, header=None)
    except Exception as e:
        print(f"   WARNING: Could not read sheet '{sheet_name}': {e}")
        return

    # get all the correct field names from our column mapping
    known_fields = list(rubric["column_mapping"].values())

    # scan every cell in the sheet
    for row_idx, row in df.iterrows():
        for col_idx, cell in enumerate(row):
            cell_str = str(cell).strip()

            # check if the cell contains a known field name
            matched_field = _match_field_name(cell_str, known_fields)

            if matched_field:
                # collect values from the same column when descending
                values = _collect_column_values(df, row_idx, col_idx)

                if values:
                    # merge with any values we already found for this field
                    existing = rubric["valid_values"].get(matched_field, [])

                    # add new values & avoid duplicates
                    combined = list(set(existing + values))
                    rubric["valid_values"][matched_field] = combined


# helper function for checking if the column value matches a known field
def _match_field_name(cell_str, known_fields):

    # normalize cell using RegEx (lowercase & and remove special characters)
    cell_normalized = re.sub(r"[^a-z0-9 ]", "", cell_str.lower()).strip()

    for field in known_fields:
        field_normalized = re.sub(r"[^a-z0-9 ]", "", field.lower()).strip()

        # check if the field name appears inside the cell text
        if field_normalized in cell_normalized: return field

    # no match
    return None


# helper function for collecting non empty values from each column starting from a row
def _collect_column_values(df, start_row, col_idx):
    values = []

    # if a cell matches a header
    field_names = [
        "extractor id", "extractor type", "trap size and units",
        "cleaning frequency", "receivingplant", "classcode",
        "secondclass", "eventtypeabbrv", "mapcategory"
    ]

    # words that are not a real value
    junk_words = [
        "total", "rows", "script", "number", "delete", "pick list",
        "description", "comments", "current", "future", "download",
        "header", "series", "interceptor", "separator", "trap -",
        "shared", "grease", "multiple", "facilities", "retired",
        "removed", "unknown", "effluent", "verified", "inactive",
        "extractor ids", "etc"
    ]

    for row_idx in range(start_row + 1, min(start_row + 50, len(df))):
        cell = str(df.iloc[row_idx, col_idx]).strip()

        # skip empty cells
        if cell in ("", "nan"):  continue

        # skip cells that are too long — real values are short
        if len(cell) > 40: continue

        # skip cells that are just a number
        if cell.isdigit():  continue
        cell_lower = cell.lower()

        # skip if the cell exactly matches a field name
        if cell_lower in field_names: continue

        # skip if the cell contains any junk words
        if any(junk in cell_lower for junk in junk_words): continue

        values.append(cell)
    return values

# helper function for building RegEx patterns based off valid values
def _build_patterns(rubric):
    for field, values in rubric["valid_values"].items():
        if not values: continue

        pattern = _infer_pattern(field, values)
        if pattern: rubric["value_patterns"][field] = pattern

# helper function for guessing the pattern
def _infer_pattern(field, values):
    # clean up the values list
    clean_values = [str(v).strip() for v in values if str(v).strip()]
    if not clean_values: return None

    # look for values that look like a extractor ID
    if "extractor id" in field.lower(): return r"^EX\s*[-]?\s*\d+"

    # if we get the exact list
    escaped = [re.escape(v) for v in clean_values]

    # build a pattern based off those reasonable values
    short_values = [v for v in escaped if len(v) <= 40]

    if short_values:
        # combine all valid values with | based off meaning or RegEx
        pattern = r"^(" + "|".join(short_values) + r")$"
        return pattern

    return None