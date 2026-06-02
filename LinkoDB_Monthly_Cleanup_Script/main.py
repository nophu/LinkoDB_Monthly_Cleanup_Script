import os
import json
import rubric_parser
import data_parser
import data_validator
from config import REPORT_CONFIG, DATA_DIR

if __name__ == '__main__':
    os.makedirs("output", exist_ok=True)

    rubric = rubric_parser.parse_rubric(os.path.join(DATA_DIR, 'Final Scripts and Update Plan.xlsx'))

    all_changes = []

    for filename, cfg in REPORT_CONFIG.items():
        print(f"\n{'='*55}")
        print(f"  Processing: {cfg['name']}")
        print(f"{'='*55}")

        filepath = os.path.join(DATA_DIR, filename)   # ← full path for reading

        records = data_parser.parse_data(filepath, rubric)

        validated, changes = data_validator.validate_data(
            records,
            rubric,
            filename,           # ← bare filename for the changes log
            only_fields=cfg["fields"]
        )

        all_changes.extend(changes)
        print(f"\n  Records validated: {len(validated)}")
        print(f"  Changes logged:    {len(changes)}")

    with open("output/all_changes.json", "w") as f:   json.dump(all_changes, f, indent=2, default=str)

    print(f"\n{'='*55}")
    print(f"  ALL DONE — {len(all_changes)} total changes")
    print(f"  Saved: output/all_changes.json")
    print(f"{'='*55}")