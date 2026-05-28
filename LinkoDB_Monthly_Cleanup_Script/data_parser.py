import re
import json
import pandas as pd

def parse_data(filepath, rubric):
    print(f"\nReading data from: {filepath}")
    wb = pd.ExcelFile(filepath, engine='openpyxl')
    sheet_name  = wb.sheet_names[0]
    print(f"Sheet: {sheet_name}")

    # read everything as raw text, no header assumed yet
    raw_df = wb.parse(sheet_name, dtype=str, header=None)
    print(f"Raw shape: {raw_df.shape}")

    # detect which file type this is
    file_type = _detect_file_type(raw_df, rubric)
    print(f"Detected file type: {file_type}")

    # parse the file based on its type
    if file_type == "permit_list":  records = _parse_permit_list(raw_df, rubric)
    elif file_type == "extract_summary": records = _parse_extract_summary(raw_df, rubric)
    elif file_type == "inspection_events": records = _parse_inspection_events(raw_df, rubric)

    # unknown structure — fall back to basic flat parse
    else:
        print("WARNING: Unknown file type — using basic flat parse")
        records = _parse_flat(raw_df, rubric)
    print(f"Total records parsed: {len(records)}")

    # save first 10 rows to JSON for inspection
    with open("output/data_parsed.json", "w") as f: json.dump(records[:10], f, indent=2, default=str)
    print("Saved: output/data_parsed.json")
    return records

def _detect_file_type(raw_df, rubric):
    messy_names = list(rubric["column_mapping"].keys())

    # convert EVERYTHING to string first before doing any stripping
    [str(v).strip() for v in raw_df.iloc[:5].values.flatten() if str(v).strip() not in ("", "nan")]

    # check row 1 specifically — inspection events has its header there
    row1_values = raw_df.iloc[1].astype(str).str.strip().tolist()

    # for inspection events
    row0_val = str(raw_df.iloc[0, 0]).strip()
    if "Fort Wayne" in row0_val and "txtPermitInfo" in row1_values: return "inspection_events"

    # for extract summaries
    all_values = raw_df.astype(str).values.flatten().tolist()
    extract_name_count = sum(1 for v in all_values if "txtExtractName" in str(v))
    if extract_name_count > 3: return "extract_summary"

    # for permit lists
    row0_values = raw_df.iloc[0].astype(str).str.strip().tolist()
    matches = sum(1 for v in row0_values if v in messy_names)
    if matches >= 1:  return "permit_list"
    return "unknown"


# parser for permit lists
def _parse_permit_list(raw_df, rubric):
    # re-read with header at row 0
    df = _reread_with_header(raw_df, 0)

    # drop empty rows and columns
    df.dropna(how="all", inplace=True)
    df.dropna(axis=1, how="all", inplace=True)

    # match and rename columns
    column_mapping, report = _match_columns(df, rubric)
    df.rename(columns=column_mapping, inplace=True)
    _print_report(report)

    # merge facility rows with extractor rows below them
    records = _merge_by_facility_column(df, facility_signal_cols=["txtPermittee", "txtPermitNo"], extractor_signal_cols=["Extractor ID", "Extractor Type", "Cleaning Frequency"])
    return records


# parser for extracted summaries
def _parse_extract_summary(raw_df, rubric):
    records     = []
    messy_names = list(rubric["column_mapping"].keys())

    current_facility  = None
    current_permit    = None
    current_col_order = None

    for row_idx in range(len(raw_df)):
        # convert every cell to string explicitly — prevents float NaN issues
        row_values = [str(v).strip() for v in raw_df.iloc[row_idx].tolist()]

        # skip completely empty rows
        non_empty = [v for v in row_values if v not in ("", "nan", "NaN")]
        if not non_empty: continue

        # detect a mini header row
        header_matches = sum(1 for v in row_values if v in messy_names)
        if header_matches >= 2:
            current_col_order = row_values
            continue

        col0  = row_values[0] if len(row_values) > 0 else ""
        col1  = row_values[1] if len(row_values) > 1 else ""
        rest_empty = all(v in ("", "nan", "NaN") for v in row_values[2:])

        # permit numbers look like DEN-0080, AG-9999, _SMV2017, HLR-147
        # must have at least one letter/underscore and one more character
        # relaxed pattern to catch all permit formats
        looks_like_permit = bool(re.match(r"^[A-Z0-9_][A-Z0-9_\-]+$", col1, re.IGNORECASE))

        if col0 not in ("", "nan", "NaN") and looks_like_permit and rest_empty:
            current_facility = col0
            current_permit   = col1
            continue

        # detect a data row
        if current_col_order is not None:
            data_values = row_values[2:]
            col_names   = current_col_order[2:]

            if any(v not in ("", "nan", "NaN") for v in data_values):
                record = {   "SiteCompany": current_facility, "PermitNo":    current_permit }

                for col_name, value in zip(col_names, data_values):
                    if col_name in ("", "nan", "NaN"): continue

                    correct_name   = rubric["column_mapping"].get(col_name, col_name)
                    # store None for empty values so JSON shows null not NaN
                    record[correct_name] = None if value in ("", "nan", "NaN") else value

                # only append if the record has at least one real value
                # besides SiteCompany and PermitNo
                real_values = {k: v for k, v in record.items()  if k not in ("SiteCompany", "PermitNo") and v is not None}
                if real_values: records.append(record)
    return records


# parser for inspection events
def _parse_inspection_events(raw_df, rubric):
    records = []

    # row 1 is the real header
    df = _reread_with_header(raw_df, 1)

    # drop empty rows and columns
    df.dropna(how="all", inplace=True)
    df.dropna(axis=1, how="all", inplace=True)

    # match and rename columns
    column_mapping, report = _match_columns(df, rubric)
    df.rename(columns=column_mapping, inplace=True)
    _print_report(report)

    # the first column contains facility info — find its name
    first_col = df.columns[0]

    current_facility_raw = None
    for _, row in df.iterrows():
        row_dict   = row.to_dict()
        first_val  = str(row_dict.get(first_col, "")).strip()

        # detect a facility row
        # facility rows start with [permit_id] like "[9199] Facility Name - Address"
        if re.match(r"^\[\d+]", first_val):
            current_facility_raw = first_val
            continue

        # detect an event row — has a ContactType/event description value
        # all other columns may have data too
        has_data = any(
            str(v).strip() not in ("", "nan")
            for k, v in row_dict.items()
            if k != first_col
        )

        if has_data and current_facility_raw:
            # parse facility name and permit id out of the raw string
            # format: "[9199] Facility Name - Address"
            facility_info = _parse_facility_string(current_facility_raw)

            record = {}
            record.update(facility_info)  # add parsed facility info
            record.update({               # add event row data
                k: (None if str(v).strip() in ("nan", "NaN") else v)
                for k, v in row_dict.items()
                if k != first_col
            })
            records.append(record)
    return records


# helper function for parsing facility string like "[9199] Café Name - 123 Main St"
def _parse_facility_string(raw_string):
    # regex to extract permit id from [9199]
    permit_match = re.match(r"^\[(\d+)]\s*(.*)", raw_string)

    if permit_match:
        permit_id    = permit_match.group(1)
        rest         = permit_match.group(2).strip()

        # split on " - " to separate name from address
        parts = rest.split("   -   ", 1)
        facility_name = parts[0].strip() if parts else rest
        address       = parts[1].strip() if len(parts) > 1 else None

        return { "PermitID":     permit_id,  "FacilityName": facility_name,  "Address":      address }

    # couldn't parse it — just return the raw string
    return {"FacilityInfo": raw_string}



# helper function for re-reading the dataframe using a specific row as the header
def _reread_with_header(raw_df, header_row_idx):
    # use the specified row as column names
    headers = raw_df.iloc[header_row_idx].astype(str).str.strip().tolist()

    # use all rows AFTER the header row as data
    data_df = raw_df.iloc[header_row_idx + 1:].copy()
    data_df.columns = headers
    data_df.reset_index(drop=True, inplace=True)
    return data_df


# helper function for merging facility rows with data rows below them
def _merge_by_facility_column(df, facility_signal_cols, extractor_signal_cols):
    records          = []
    current_facility = {}

    for _, row in df.iterrows():
        row_dict = row.to_dict()

        # check if this row has facility info
        has_facility = any(
            str(row_dict.get(col, "")).strip() not in ("", "nan")
            for col in facility_signal_cols
            if col in row_dict
        )

        # check if this row has extractor/data info
        has_data = any(
            str(row_dict.get(col, "")).strip() not in ("", "nan")
            for col in extractor_signal_cols
            if col in row_dict
        )

        if has_facility:
            # save facility info to attach to rows below
            current_facility = {
                col: row_dict.get(col)
                for col in facility_signal_cols
                if col in row_dict
            }

        if has_data:
            # combine facility info + data row into one record
            combined = {}
            combined.update(current_facility)
            combined.update({
                k: v for k, v in row_dict.items()
                if k not in facility_signal_cols
            })

            # clean up NaN values
            cleaned = {
                k: (None if str(v).strip() in ("nan", "NaN") else v)
                for k, v in combined.items()
            }

            records.append(cleaned)
    return records


# helper function for basic flat parse fallback for unknown file types
def _parse_flat(raw_df, rubric):
    header_row = _find_header_row(raw_df, rubric)
    df         = _reread_with_header(raw_df, header_row)

    df.dropna(how="all", inplace=True)
    df.dropna(axis=1, how="all", inplace=True)

    column_mapping, report = _match_columns(df, rubric)
    df.rename(columns=column_mapping, inplace=True)
    _print_report(report)
    return df.to_dict(orient="records")


# helper function for finding ind the real header row by matching against rubric names
def _find_header_row(raw_df, rubric):
    messy_names = list(rubric["column_mapping"].keys())
    best_row    = 0
    best_score  = 0

    for row_idx in range(min(20, len(raw_df))):
        row_values = raw_df.iloc[row_idx].astype(str).str.strip().tolist()
        score      = sum(1 for v in row_values if v in messy_names)

        if score > best_score:
            best_score = score
            best_row   = row_idx
    return best_row


# helper function for matching each column to a rubric field
def _match_columns(df, rubric):
    report = {
        "matched":       [],
        "auto_fixed":    [],
        "not_in_rubric": [],
        "flagged":       [],
        "ignored":       []
    }

    column_mapping = {}
    for col in df.columns:
        col_str = str(col).strip()

        # pass 1 — match by name
        matched = _match_by_name(col_str, rubric["column_mapping"])
        if matched:
            column_mapping[col] = matched
            report["matched"].append({"original": col, "renamed_to": matched})
            continue

        # pass 2 — match by value pattern
        values  = df[col].dropna().astype(str).tolist()
        matched = _match_by_values(values, rubric["value_patterns"])
        if matched:
            column_mapping[col] = matched
            report["auto_fixed"].append({"original": col, "renamed_to": matched})
            continue

        # no match
        if _looks_important(col_str): report["not_in_rubric"].append({"column": col})
        else:  report["ignored"].append(col)
    return column_mapping, report


# helper function for matching column name against rubric mapping
def _match_by_name(col, column_mapping):
    if col in column_mapping: return column_mapping[col]

    col_norm = re.sub(r"[^a-z0-9]", "", col.lower())
    for messy, correct in column_mapping.items():
        if col_norm == re.sub(r"[^a-z0-9]", "", messy.lower()):  return correct
    return None


# helper function for matching column by sampling values against rubric patterns
def _match_by_values(values, value_patterns):
    sample = [v.strip() for v in values if v.strip() not in ("", "nan")][:20]

    if not sample:  return None

    best_field = None
    best_score = 0.0

    for field, pattern in value_patterns.items():
        matches = sum(1 for v in sample if re.match(pattern, v, re.IGNORECASE))
        score   = matches / len(sample)

        if score > best_score:
            best_score = score
            best_field = field

    if best_score >= 0.60: return best_field
    return None


# helper function for deciding if an unmatched column looks important
def _looks_important(col):
    generic = re.match(r"^(text|label|field|col|column)\d*$", col.lower())
    return not generic


# helper for print column matching report to console
def _print_report(report):
    print("\n" + "=" * 45)
    print("  COLUMN MATCHING REPORT")
    print("=" * 45)

    print(f"\n  Matched by name:      {len(report['matched'])}")
    for r in report["matched"]:  print(f"      '{r['original']}'  :  '{r['renamed_to']}'")

    print(f"\n  Matched by values:    {len(report['auto_fixed'])}")
    for r in report["auto_fixed"]:   print(f"      '{r['original']}'  :  '{r['renamed_to']}'")

    print(f"\n  Not in rubric:        {len(report['not_in_rubric'])} columns")
    for r in report["not_in_rubric"]:  print(f"      '{r['column']}'")

    print(f"\n  Ignored:              {len(report['ignored'])} columns")
    for col in report["ignored"]:  print(f"      '{col}'")
    print("\n" + "=" * 45 + "\n")