import re

def validate_data(records, rubric, source_filename):
    print(f"\nValidating {len(records)} records from {source_filename}...")

    # only validate fields the rubric has rules for
    checkable_fields = list(rubric["valid_values"].keys())
    print(f"   Fields being validated: {checkable_fields}")

    validated  = []  # cleaned records
    changes    = []  # log of every change or flag

    for record in records:
        cleaned_record = {}

        for field, value in record.items():

            # skip empty values — nothing to validate
            if value is None or str(value).strip() in ("", "nan", "NaN"):
                cleaned_record[field] = value
                continue

            value_str = str(value).strip()

            # only validate fields the rubric has rules for
            # everything else passes through unchanged
            if field not in checkable_fields:
                cleaned_record[field] = value
                continue

            # run the value through our checks
            result = _check_value(field, value_str, rubric)

            # store the cleaned value
            cleaned_record[field] = result["cleaned_value"]

            # log anything that wasn't a clean pass
            if result["status"] != "pass":
                changes.append({
                    # where it came from
                    "source_file":   source_filename,
                    "facility":      _get_facility_name(record),
                    "permit_no":     _get_permit_no(record),

                    # what field and what changed
                    "field":         field,
                    "original":      value_str,
                    "cleaned_value": result["cleaned_value"],
                    "status":        result["status"],
                    "note":          result["note"]
                })
        validated.append(cleaned_record)

    # print a summary to the console
    _print_summary(changes, source_filename)
    return validated, changes


# helper function for checking a single value against the rubric rules for its field
def _check_value(field, value, rubric):
    valid_values  = rubric["valid_values"].get(field, [])
    value_pattern = rubric["value_patterns"].get(field)

    # exact match
    if value in valid_values: return {"status":        "pass", "cleaned_value": value, "note":          "exact match"}

    # case-insensitive match
    value_lower = value.lower()
    for valid in valid_values:
        if valid.lower() == value_lower: return { "status":        "fixed", "cleaned_value": valid, "note":          f"fixed casing: '{value}' is now: '{valid}'" }

    # regex pattern match
    if value_pattern:
        if re.match(value_pattern, value, re.IGNORECASE): return { "status":        "pass", "cleaned_value": value, "note":          "matched pattern" }

    # partial match
    close = _find_partial_match(value, valid_values)
    if close: return {"status":        "flagged", "cleaned_value": value,   "note":          f"close to '{close}' but not exact — review manually"}

    # no match at all
    return { "status":        "flagged", "cleaned_value": value, "note":          f"'{value}' not in rubric valid values for '{field}' — review manually" }


# helper function for finding a partial match between a value and the valid values
def _find_partial_match(value, valid_values):
    value_lower = value.lower()

    for valid in valid_values:
        valid_lower = valid.lower()

        # check if one string contains the other
        if value_lower in valid_lower or valid_lower in value_lower: return valid
    return None

# helper function for getting the facility name from a record
def _get_facility_name(record):
    for field in ["txtPermittee", "SiteCompany", "FacilityName", "Permittee"]:
        val = record.get(field)
        if val and str(val).strip() not in ("", "nan", "NaN", "None"): return str(val).strip()
    return "Unknown"

# helper function for getting the permit number from a record
def _get_permit_no(record):
    for field in ["txtPermitNo", "PermitNo", "PermitID", "PermitNo"]:
        val = record.get(field)
        if val and str(val).strip() not in ("", "nan", "NaN", "None"):  return str(val).strip()
    return "Unknown"

# helper function for printing a readable summary to the console
def _print_summary(changes, source_filename):
    fixed   = [c for c in changes if c["status"] == "fixed"]
    flagged = [c for c in changes if c["status"] == "flagged"]

    print("\n" + "=" * 45)
    print(f"  VALIDATION SUMMARY — {source_filename}")
    print("=" * 45)
    print(f"\n  Total changes logged: {len(changes)}")
    print(f"  Auto fixed:           {len(fixed)}")
    print(f"  Flagged for review:   {len(flagged)}")

    if fixed:
        print(f"\n  AUTO FIXED:")
        for c in fixed:
            print(f"      [{c['facility']} | {c['permit_no']}]")
            print(f"      {c['field']}: '{c['original']}' to '{c['cleaned_value']}'")

    if flagged:
        print(f"\n  FLAGGED FOR REVIEW:")
        for c in flagged[:20]:  # show first 20 only so console doesn't flood
            print(f"      [{c['facility']} | {c['permit_no']}]")
            print(f"      {c['field']}: '{c['original']}' — {c['note']}")
        if len(flagged) > 20: print(f"      ... and {len(flagged) - 20} more — see output/changes.json")

    print("\n" + "=" * 45 + "\n")