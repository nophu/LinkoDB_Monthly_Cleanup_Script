import re
import json

# Known value equivalents that map to a canonical valid value (auto-fixed).
# e.g. "Twice Monthly" means the same as "Every 2 Weeks" per the program.
FIELD_ALIASES = {
    "Cleaning Frequency": {
        "twice monthly": "Every 2 Weeks",
    },
}

# Old-scheme extractor migration. When an old ID (number 1-99) has a description
# containing one of these keywords, map the number into the matching new range:
#     new number = base + (old_number - 1)
# e.g. a "grease trap" with ID 001 -> EX200, 002 -> EX201, ... up to range_max.
EXTRACTOR_MIGRATION = [
    # unambiguous type words in the description → new series (confident, GREEN)
    {"keyword": "grease trap",    "base": 200, "range_max": 220, "label": "grease trap"},
    {"keyword": "sediment",       "base": 230, "range_max": 239, "label": "sediment trap"},
    {"keyword": "hair",           "base": 250, "range_max": 259, "label": "hair trap"},
    {"keyword": "amalgam",        "base": 300, "range_max": 310, "label": "amalgam separator"},
    {"keyword": "oil/sand",       "base": 311, "range_max": 319, "label": "oil/sand separator"},
    {"keyword": "oil water",      "base": 311, "range_max": 319, "label": "oil/sand separator"},
    {"keyword": "oil-water",      "base": 311, "range_max": 319, "label": "oil/sand separator"},
    {"keyword": "sand separator", "base": 311, "range_max": 319, "label": "oil/sand separator"},
    {"keyword": "solids",         "base": 150, "range_max": 159, "label": "solids interceptor"},
    {"keyword": "shared",         "base": 160, "range_max": 169, "label": "shared interceptor"},
]

# Grease interceptors split by material. Only used once we know it's an interceptor.
# Concrete → EX100 series, everything else → EX120 series.
INTERCEPTOR_MATERIAL = [
    ("concrete",   100, 110, "grease interceptor (concrete)"),
    ("fiberglass", 120, 129, "grease interceptor (fiberglass)"),
    ("plastic",    120, 129, "grease interceptor (plastic)"),
    ("steel",      120, 129, "grease interceptor (steel)"),
]

# Fallback when the data description names no type: the OLD code number itself
# identifies the type (per the old rubric). These produce YELLOW "suggested" fixes
# for a person to confirm — not silent auto-fixes.
OLD_CODE_FALLBACK = {
    50: (160, "shared interceptor"),
    60: (150, "solids interceptor"),
    75: (311, "oil/sand separator"),
    80: (400, "food-truck effluent valve"),
    98: (500, "retired/inactive"),
    99: (600, "no interceptor (verified)"),
}

# fields that may hold the extractor's free-text description, across report types
_DESC_FIELDS = ("ExtractDesc", "Text24", "Description", "ExtractorDesc")

def _get_description(record):
    if not record:
        return ""
    for f in _DESC_FIELDS:
        v = record.get(f)
        if v and str(v).strip().lower() not in ("", "nan", "none"):
            return str(v).strip()
    return ""

def validate_data(records, rubric, source_filename, only_fields=None):
    print(f"\nValidating {len(records)} records from {source_filename}...")

    # only validate fields the rubric has rules for
    checkable_fields = list(rubric["valid_values"].keys())

    # if only_fields is provided, restrict to just those fields
    # this lets each report only check the fields that apply to it
    if only_fields is not None:
        checkable_fields = [f for f in checkable_fields if f in only_fields]

    print(f"   Fields being validated: {checkable_fields}")

    validated = []   # cleaned records
    changes   = []   # log of every change or flag

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
            result = _check_value(field, value_str, rubric, record)

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
                    "note":          result["note"],
                })

        validated.append(cleaned_record)

    # save changes to JSON so the summary report can use them
    _save_changes(changes, source_filename)

    # print a summary to the console
    _print_summary(changes, source_filename)
    return validated, changes


# helper function for checking a single value against the rubric rules for its field
def _check_value(field, value, rubric, record=None):
    valid_values  = rubric["valid_values"].get(field, [])
    value_pattern = rubric["value_patterns"].get(field)

    # special rule: Trap Size and Units uses a size-based unit check, not a static list
    # rule: size <= 99 → units must be gpm | size >= 100 → units must be gal
    if field == 'Trap Size and Units':
        return _check_trap_size(value)

    # special rule: Extractor IDs must start with "EX" per the rubric
    # a bare number like "001" / "050" just needs the EX prefix → "EX001", "EX050"
    if field == 'Extractor ID':
        return _check_extractor_id(value, valid_values, value_pattern, _get_description(record))

    # known equivalents → auto-map to the canonical valid value
    alias = FIELD_ALIASES.get(field, {}).get(value.lower())
    if alias:
        return {
            "status":        "fixed",
            "cleaned_value": alias,
            "note":          f"mapped '{value}' → '{alias}'",
        }

    # special case: empty valid list means the field should be blank
    # (e.g. TrunkLine — rubric says "delete entry and leave field BLANK")
    if valid_values == []:
        return {
            "status":        "flagged",
            "cleaned_value": value,
            "note":          f"'{value}' should be deleted — leave this field blank per rubric",
        }

    # exact match — value is already correct
    if value in valid_values:
        return {"status": "pass", "cleaned_value": value, "note": "exact match"}

    # case-insensitive match — value is right but wrong casing, auto-fix it
    value_lower = value.lower()
    for valid in valid_values:
        if valid.lower() == value_lower:
            return {
                "status":        "fixed",
                "cleaned_value": valid,
                "note":          f"fixed casing: '{value}' → '{valid}'",
            }

    # regex pattern match — value matches the expected format
    if value_pattern:
        if re.match(value_pattern, value, re.IGNORECASE):
            return {"status": "pass", "cleaned_value": value, "note": "matched pattern"}

    # partial match — value is close to something valid, flag for manual review
    close = _find_partial_match(value, valid_values)
    if close:
        return {
            "status":        "flagged",
            "cleaned_value": value,
            "note":          f"'{value}' is close to '{close}' — should it be changed to '{close}'?",
        }

    # no match at all — flag for manual review
    return {
        "status":        "flagged",
        "cleaned_value": value,
        "note":          f"'{value}' is not a valid value for '{field}' — review manually",
    }


# special rule for Trap Size and Units — not a static list, it's a size-based unit check
# rubric rule: size <= 99 → gpm | size >= 100 → gal | blank → leave blank
def _check_trap_size(value):
    parts = value.strip().split()

    # must be exactly two parts: a number and a unit
    if len(parts) != 2:
        return {
            "status":        "flagged",
            "cleaned_value": value,
            "note":          f"expected format '<number> <unit>' (e.g. '35 gpm') — got '{value}'",
        }

    size_str, unit = parts[0], parts[1]

    # size must be numeric
    try:
        size = float(size_str)
    except ValueError:
        return {
            "status":        "flagged",
            "cleaned_value": value,
            "note":          f"'{size_str}' is not a valid numeric trap size",
        }

    correct_unit = "gpm" if size <= 99 else "gal"

    # unit is already correct
    if unit == correct_unit:
        return {"status": "pass", "cleaned_value": value, "note": "exact match"}

    # unit is correct but wrong casing — auto-fix
    if unit.lower() == correct_unit:
        fixed = f"{size_str} {correct_unit}"
        return {
            "status":        "fixed",
            "cleaned_value": fixed,
            "note":          f"fixed casing: '{value}' → '{fixed}'",
        }

    # unit is wrong — but the rule is deterministic (size decides the unit),
    # so this is a certain correction → auto-fix
    fixed = f"{size_str} {correct_unit}"
    return {
        "status":        "fixed",
        "cleaned_value": fixed,
        "note":          f"corrected unit by rule (≤99 → gpm, ≥100 → gal): '{value}' → '{fixed}'",
    }


# special rule for Extractor IDs — per the rubric every ID must start with "EX"
# AND fall within a valid NEW range (EX100-EX830). Bare numbers get "EX" added,
# then we check the number against the new ranges. Old-scheme IDs (EX001-099) and
# non-numeric values can't be auto-fixed → manual review.
def _check_extractor_id(value, valid_values, value_pattern, description=""):
    # parse the valid NEW ranges from values like "EX100 - EX110" or "EX400"
    ranges = []
    for v in valid_values:
        nums = re.findall(r"\d+", v)
        if len(nums) >= 2:
            ranges.append((int(nums[0]), int(nums[1])))
        elif len(nums) == 1:
            ranges.append((int(nums[0]), int(nums[0])))

    raw = value.strip()

    # figure out the candidate EX-id
    if re.match(r"^\d+$", raw):
        candidate   = "EX" + raw            # bare number → add prefix
    elif re.match(r"^EX\s*\d+$", raw, re.IGNORECASE):
        candidate   = "EX" + re.sub(r"[^0-9]", "", raw)   # normalise existing EX id
    else:
        # non-numeric like "HSW - Station 1" — can't auto-handle
        return {
            "status":        "flagged",
            "cleaned_value": value,
            "note":          f"'{value}' is not a standard extractor ID — review manually",
        }

    num = int(re.sub(r"[^0-9]", "", candidate))
    in_new_range = any(lo <= num <= hi for lo, hi in ranges)

    if in_new_range:
        if candidate == raw:
            return {"status": "pass", "cleaned_value": candidate, "note": "valid"}
        return {
            "status":        "fixed",
            "cleaned_value": candidate,
            "note":          f"added EX prefix: '{value}' → '{candidate}'",
        }

    # number is not in any NEW range
    if num < 100:
        desc = (description or "").lower()

        # TIER 1 — description names a clear trap/separator type → confident (GREEN)
        for m in EXTRACTOR_MIGRATION:
            if m["keyword"] in desc:
                new_num = m["base"] + (num - 1)
                if m["base"] <= new_num <= m["range_max"]:
                    new_id = f"EX{new_num}"
                    return {
                        "status":        "fixed",
                        "cleaned_value": new_id,
                        "note":          f"migrated old {m['label']} ID '{value}' → '{new_id}'",
                    }

        # interceptor: must know the material to pick concrete (EX100) vs other (EX120)
        if "interceptor" in desc:
            for kw, base, cap, label in INTERCEPTOR_MATERIAL:
                if kw in desc:
                    new_num = base + (num - 1)
                    if base <= new_num <= cap:
                        new_id = f"EX{new_num}"
                        return {
                            "status":        "fixed",
                            "cleaned_value": new_id,
                            "note":          f"migrated old {label} ID '{value}' → '{new_id}'",
                        }
            # interceptor but material not stated → too unspecific to pick → review
            return {
                "status":        "flagged",
                "cleaned_value": value,
                "note":          f"'{value}' is an old grease interceptor but the material isn't stated — assign EX100 (concrete) or EX120 (other) manually",
            }

        # TIER 2 — no type in description, but the OLD code number implies the type
        #          (per the old rubric) → suggest the new ID for review (YELLOW)
        fb = OLD_CODE_FALLBACK.get(num)
        if fb:
            base, label = fb
            new_id = f"EX{base}"
            return {
                "status":        "flagged",
                "cleaned_value": value,
                "note":          f"old code '{candidate}' is a {label} per the old rubric — suggested: '{new_id}'",
            }

        # TIER 3 — plain sequence number (EX001-EX049) with no type stated:
        #          can't tell trap from interceptor → review manually
        return {
            "status":        "flagged",
            "cleaned_value": value,
            "note":          f"'{value}' is an old-scheme ID ({candidate}) with no type stated — assign a new EX ID manually",
        }
    # at or above EX100 but not inside any valid range → just invalid
    return {
        "status":        "flagged",
        "cleaned_value": value,
        "note":          f"'{value}' ({candidate}) is not in any valid range (EX100–EX830) — review manually",
    }


# helper function for finding a partial match between a value and the valid values
def _find_partial_match(value, valid_values):
    value_lower = value.lower()
    for valid in valid_values:
        valid_lower = valid.lower()

        # data value starts with a valid value
        # e.g. "Multi-Tenant Facility - LARGE" starts with "Multi-Tenant"
        if value_lower.startswith(valid_lower):
            return valid

        # valid value starts with the data value
        # e.g. data "Mobile" vs valid "Mobile Business"
        if valid_lower.startswith(value_lower):
            return valid

    return None


# helper function for getting the facility name from a record
# checks several common field name variations across different file types
def _get_facility_name(record):
    for field in ["txtPermittee", "SiteCompany", "FacilityName", "Permittee",
                  "PermitteeAccount", "AccountName", "Name", "FacilityInfo"]:
        val = record.get(field)
        if val and str(val).strip() not in ("", "nan", "NaN", "None"):
            return str(val).strip()
    return "Unknown"


# helper function for getting the permit number from a record
def _get_permit_no(record):
    for field in ["txtPermitNo", "PermitNo", "PermitID", "PermitNumber", "Permit"]:
        val = record.get(field)
        if val and str(val).strip() not in ("", "nan", "NaN", "None"):
            return str(val).strip()
    return "Unknown"


# saves changes to a JSON file named after the source file
def _save_changes(changes, source_filename):
    # strip the extension and use it as the output filename
    base = source_filename.replace(".xlsx", "").replace(".csv", "")
    output_path = f"output/{base}_changes.json"
    with open(output_path, "w") as f:
        json.dump(changes, f, indent=2, default=str)
    print(f"   Saved: {output_path}")


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
            print(f"      {c['field']}: {c['note']}")

    if flagged:
        print(f"\n  FLAGGED FOR REVIEW:")
        for c in flagged:
            print(f"      [{c['facility']} | {c['permit_no']}]")
            print(f"      {c['field']}: {c['note']}")

    print("\n" + "=" * 45 + "\n")