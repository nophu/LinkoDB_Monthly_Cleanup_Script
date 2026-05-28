import re
import json
import pandas as pd

def parse_data(filepath, rubric):
    print(f"\nReading data from: {filepath}")
    wb = pd.ExcelFile(filepath, engine='openpyxl')
    sheet_names = wb.sheet_names
    print(f"Sheet names: {sheet_names}")

    # only look at the first sheet
    sheet_name = sheet_names[0]

    # read everything as text, no header assumed yet
    df = wb.parse(sheet_name, dtype=str, header=None)

    # find which row is the real header row
    header_row = _find_header_row(df, rubric)
    print(f"Header row found at row index: {header_row}")

    # re-read the file using the correct header row
    df = wb.parse(sheet_name, dtype=str, header=header_row, engine='openpyxl')

    # strip whitespace from column names
    df.columns = df.columns.astype(str).str.strip()

    # drop completely empty columns
    df.dropna(axis=1, how="all", inplace=True)

    # drop completely empty rows
    df.dropna(how="all", inplace=True)

    print(f"Columns found: {df.columns.tolist()}")
    print(f"Rows found: {len(df)}")

    # match each column to a rubric field
    column_mapping, report = _match_columns(df, rubric)

    # rename columns using the mapping we built
    df.rename(columns=column_mapping, inplace=True)

    # print the matching report to console
    _print_report(report)

    # merge facility rows with their extractor rows
    # so each record has both facility info AND extractor info
    records = _merge_facility_and_extractor_rows(df)

    # save first 10 rows to json for inspection
    with open("output/data_parsed.json", "w") as f: json.dump(records[:10], f, indent=2, default=str)
    print("Saved first 10 rows to output/data_parsed.json")
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


# helper function for matching each column to a rubric field
def _match_columns(df, rubric):
    # matched by name
    # matched by value pattern
    # column exists in data but rubric doesn't mention it
    # something looks wrong — needs human review
    # generic placeholder columns like Text50, Label130
    report = {  "matched":       [],    "auto_fixed":    [],   "not_in_rubric": [],   "flagged":       [],   "ignored":       []   }

    # old name → new name
    column_mapping = {}

    for col in df.columns:
        col_str = str(col).strip()

        # match by column name
        matched = _match_by_name(col_str, rubric["column_mapping"])

        if matched:
            column_mapping[col] = matched
            report["matched"].append({ "original":   col,   "renamed_to": matched, "method":     "name_match" })
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
        if _looks_important(col_str):
            # has a real name but rubric doesn't know about it — ok to keep
            report["not_in_rubric"].append({ "column": col,  "note":   "Not in rubric — may be extra data column, ok to keep as is"})
        else:
            # generic placeholder name — safe to ignore
            report["ignored"].append(col)

    return column_mapping, report


# merge facility rows with their extractor rows
def _merge_facility_and_extractor_rows(df):
    merged_records = []

    # columns that tell us a row is a facility row
    facility_cols = ['txtPermittee', 'txtPermitNo', 'txtSiteAddr1']

    # columns that tell us a row is an extractor row
    # we use the RENAMED versions here since we already renamed the columns
    extractor_cols = ['Extractor ID', 'Extractor Type', 'Cleaning Frequency']

    # keep track of the last facility we saw
    current_facility = {}

    for _, row in df.iterrows():
        row_dict = row.to_dict()

        # check if this row has facility info
        has_facility = any(
            str(row_dict.get(col, '')).strip() not in ('', 'nan', 'NaN')
            for col in facility_cols
            if col in row_dict
        )

        # check if this row has extractor info
        has_extractor = any(
            str(row_dict.get(col, '')).strip() not in ('', 'nan', 'NaN')
            for col in extractor_cols
            if col in row_dict
        )

        if has_facility:
            # save the facility info to attach to extractor rows below
            current_facility = {
                col: row_dict.get(col)
                for col in facility_cols
                if col in row_dict
            }

        if has_extractor:
            # combine facility info + extractor info into one record
            combined = {}
            combined.update(current_facility)   # facility info first
            combined.update({                    # then extractor info
                k: v for k, v in row_dict.items()
                if k not in facility_cols        # don't overwrite facility info
            })

            # replace NaN with None so JSON looks clean
            cleaned = {
                k: (None if str(v).strip() in ('nan', 'NaN') else v)
                for k, v in combined.items()
            }

            merged_records.append(cleaned)

    print(f"Merged into {len(merged_records)} combined records")
    return merged_records


# helper function for matching a column name against the rubric column_mapping
def _match_by_name(col, column_mapping):
    # exact match
    if col in column_mapping: return column_mapping[col]

    # normalized match — strip everything except letters and numbers
    # so "CleaningFreq" and "cleaningfreq" both match
    col_norm = re.sub(r"[^a-z0-9]", "", col.lower())

    for messy, correct in column_mapping.items():
        messy_norm = re.sub(r"[^a-z0-9]", "", messy.lower())
        if col_norm == messy_norm: return correct

    return None


# helper function or matching a column by sampling its values against rubric patterns
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
    if best_score >= 0.60:   return best_field
    return None


# helper function for deciding if an unmatched column looks important enough to flag
def _looks_important(col):
    # if it matches a generic placeholder pattern, ignore it
    generic = re.match(r"^(text|label|field|col|column)\d*$", col.lower())
    if generic: return False
    return True


# helper function for printing a human-readable summary to the console
def _print_report(report):
    print("\n" + "=" * 45)
    print("  COLUMN MATCHING REPORT")
    print("=" * 45)

    print(f"\n  Matched by name:      {len(report['matched'])}")
    for r in report["matched"]: print(f"      '{r['original']}'  →  '{r['renamed_to']}'")

    print(f"\n  Matched by values:    {len(report['auto_fixed'])}")
    for r in report["auto_fixed"]:   print(f"      '{r['original']}'  →  '{r['renamed_to']}'")

    print(f"\n  Flagged:              {len(report['flagged'])}")
    for r in report["flagged"]:   print(f"      '{r['column']}' — {r['note']}")

    print(f"\n  Not in rubric:        {len(report['not_in_rubric'])} columns")
    for col in report["not_in_rubric"]: print(f"      '{col['column']}' — {col['note']}")

    print(f"\n  Ignored:              {len(report['ignored'])} columns")
    for col in report["ignored"]: print(f"      '{col}'")

    print("\n" + "=" * 45 + "\n")