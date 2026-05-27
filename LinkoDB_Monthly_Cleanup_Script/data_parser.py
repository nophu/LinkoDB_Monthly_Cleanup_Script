import re
import json
import pandas as pd

def parse_data(filepath, rubric):
    print(f"\nReading data from: {filepath}")
    wb = pd.ExcelFile(filepath)
    sheet_names = wb.sheet_names
    print(f"Sheet names: {sheet_names}")

    # check only the first sheet
    sheet_name = sheet_names[0]

    # read everything as text, assuming there exists NO header
    df = wb.parse(sheet_name, dtype=str, header=None)

   # find which row is the real header row
    header_row = _find_header_row(df, rubric)
    print(f"Header row found at row index: {header_row}")

    # reread file
    df = wb.parse(sheet_name, dtype=str, header=header_row)

    # strip white spaces
    df.columns = df.columns.astype(str).str.strip()

    # drop empty columns
    df.dropna(axis=1, how="all", inplace=True)
    print(f" Columns found: {df.columns.tolist()}")
    print(f"  Rows found: {len(df)}")

    # match each column to a rubric field
    column_mapping, report = _match_columns(df, rubric)

    # rename columns
    df.rename(columns=column_mapping, inplace=True)

    # print the report in console
    _print_report(report)

    # save to json
    records = df.to_dict(orient="records")
    with open("output/data_parsed.json", "w") as f: json.dump(records[10:], f, indent=2, default=str)
    print(" Saved first 10 rows to output/data_parsed.json")
    return df, report

# helper function for finding the real header row
def _find_header_row(df, rubric):
    messy_names = list(rubric["column_mapping"].keys())
    best_row    = 0
    best_score  = 0

    # only check the first 20 rows — header is never far down
    for row_idx in range(min(20, len(df))):
        row_values = df.iloc[row_idx].astype(str).str.strip().tolist()

        # count how many values in this row match a known messy column name
        score = sum(1 for v in row_values if v in messy_names)

        if score > best_score:
            best_score = score
            best_row   = row_idx

    return best_row

# match each column to a rubric field
def _match_columns(df, rubric):
    # matched by name
    # matched by value pattern
    # could not match
    # column not in rubric at all, probably ok to skip
    report = { "matched":    [],   "auto_fixed": [], "not_in_rubric": [],  "flagged":    [],   "ignored":    []}

    # old name to new name
    column_mapping = {}
    for col in df.columns:
        col_str = str(col).strip()

        # match by column name
        matched = _match_by_name(col_str, rubric["column_mapping"])

        if matched:
            column_mapping[col] = matched
            report["matched"].append({
                "original":   col,
                "renamed_to": matched,
                "method":     "name_match"
            })
            continue

        # match by value pattern
        values = df[col].dropna().astype(str).tolist()
        matched = _match_by_values(values, rubric["value_patterns"])

        if matched:
            column_mapping[col] = matched
            report["auto_fixed"].append({
                "original":   col,
                "renamed_to": matched,
                "method":     "value_match"
            })
            continue

        # no match found
        # check if it looks important or just extra
        if _looks_important(col_str): report["not_in_rubric"].append({ "column": col, "note":   "Could not match to rubric - may be extra data column, ok to keep as is" })
        else: report["ignored"].append(col)
    return column_mapping, report

# helper function for matching a column name against our column_mapping
def _match_by_name(col, column_mapping):
    # exact match
    if col in column_mapping: return column_mapping[col]

    # normalized match, so we strip EVERYTHING except letters and numbers
    col_norm = re.sub(r"[^a-z0-9]", "", col.lower())

    for messy, correct in column_mapping.items():
        messy_norm = re.sub(r"[^a-z0-9]", "", messy.lower())
        if col_norm == messy_norm: return correct
    return None

# helper function for matching a column by sampling its values against rubric patterns
def _match_by_values(values, value_patterns):
    # take a sample of up to 20 non-empty values
    sample = [v.strip() for v in values if v.strip() not in ("", "nan")][:20]

    if not sample: return None
    best_field = None
    best_score = 0.0

    for field, pattern in value_patterns.items():
        # count how many sample values match this pattern
        matches = sum(1 for v in sample if re.match(pattern, v, re.IGNORECASE))
        score   = matches / len(sample)

        if score > best_score:
            best_score = score
            best_field = field

    # only return a match if at least 60% of values matched
    if best_score >= 0.60: return best_field
    return None

# helper function for deciding if an unmatched column looks important, we flag it
def _looks_important(col):
    # if it looks like a generic placeholder name, ignore it
    generic = re.match(r"^(text|label|field|col|column)\d*$", col.lower())
    if generic: return False
    return True

# if you want a readable summary, we can print it to the console
def _print_report(report):
    print("\n" + "=" * 45)
    print("  COLUMN MATCHING REPORT")
    print("=" * 45)

    print(f"\n  Matched by name:    {len(report['matched'])}")
    for r in report["matched"]: print(f"      '{r['original']}'  →  '{r['renamed_to']}'")

    print(f"\n  Matched by values:  {len(report['auto_fixed'])}")
    for r in report["auto_fixed"]:  print(f"      '{r['original']}'  →  '{r['renamed_to']}'")

    print(f"\n   Flagged:            {len(report['flagged'])}")
    for r in report["flagged"]:  print(f"      '{r['column']}' — {r['note']}")

    print(f"\n   Not in rubric:            {len(report['not_in_rubric'])} columns")
    for col in report["not_in_rubric"]:  print(f"      '{col}'")

    print("\n" + "=" * 45 + "\n")
