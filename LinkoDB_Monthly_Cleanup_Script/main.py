import os
import json
import rubric_parser
import data_parser
import data_validator
from config import REPORT_CONFIG

if __name__ == '__main__':
    os.makedirs("output", exist_ok=True)

    # parse the rubric once — all files share the same rules
    rubric = rubric_parser.parse_rubric('Final Scripts and Update Plan.xlsx')

    all_changes = []  # collect changes from every file for the summary report

    # run every file through the parser and validator
    for filename, cfg in REPORT_CONFIG.items():
        print(f"\n{'='*55}")
        print(f"  Processing: {cfg['name']}")
        print(f"{'='*55}")

        # parse the file into records
        records = data_parser.parse_data(filename, rubric)

        # validate only the fields that apply to this report
        validated, changes = data_validator.validate_data(
            records,
            rubric,
            filename,
            only_fields=cfg["fields"]   # per-report field restriction
        )

        all_changes.extend(changes)
        print(f"\n  Records validated: {len(validated)}")
        print(f"  Changes logged:    {len(changes)}")

    # save one combined changes file covering all reports
    with open("output/all_changes.json", "w") as f:  json.dump(all_changes, f, indent=2, default=str)

    print(f"\n{'='*55}")
    print(f"  ALL DONE")
    print(f"  Total changes across all files: {len(all_changes)}")
    print(f"  Saved: output/all_changes.json")
    print(f"{'='*55}")