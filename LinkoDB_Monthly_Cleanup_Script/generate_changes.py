import os
import json
import rubric_parser
import data_parser
import data_validator
import summary_report
from config import REPORT_CONFIG, match_report


def pick_files():
    """Open native file-explorer dialogs to choose the rubric and the data files."""
    import tkinter as tk
    from tkinter import filedialog, messagebox

    root = tk.Tk()
    root.withdraw()                      # hide the empty tk window
    root.attributes("-topmost", True)    # keep dialogs in front
    xlsx = [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]

    messagebox.showinfo("Step 1 of 2", "Select the RUBRIC file\n(Final Scripts and Update Plan).")
    rubric_path = filedialog.askopenfilename(title="Select the RUBRIC file", filetypes=xlsx)
    if not rubric_path: return None, []

    messagebox.showinfo("Step 2 of 2",
                        "Select the report file(s) to check.\n\n"
                        "Hold Ctrl (or Shift) to select more than one.")
    data_paths = filedialog.askopenfilenames(title="Select the report file(s) to check",     filetypes=xlsx)
    root.destroy()
    return rubric_path, list(data_paths)


if __name__ == "__main__":
    os.makedirs("output", exist_ok=True)

    rubric_path, data_paths = pick_files()
    if not rubric_path:
        print("No rubric selected — nothing to do.")
        raise SystemExit
    if not data_paths:
        print("No report files selected — nothing to do.")
        raise SystemExit

    rubric = rubric_parser.parse_rubric(rubric_path)
    all_changes = []
    processed   = []   # which report keys we actually ran (for the summary)

    for path in data_paths:
        filename = os.path.basename(path)
        key, cfg = match_report(filename)

        if cfg is None:
            print(f"\n  SKIPPED: '{filename}' doesn't match any known report.")
            print(f"  Known reports: {', '.join(c['name'] for c in REPORT_CONFIG.values())}")
            continue

        print(f"\n{'='*55}\n  Processing: {cfg['name']}\n{'='*55}")
        records = data_parser.parse_data(path, rubric)

        # use the stable report KEY as the identifier so the summary lines up
        _, changes = data_validator.validate_data(
            records, rubric, key, only_fields=cfg["fields"]
        )
        all_changes.extend(changes)
        processed.append(key)

    with open("output/all_changes.json", "w") as f:  json.dump(all_changes, f, indent=2, default=str)

    print(f"\n{'='*55}")
    print(f"  Total changes across {len(processed)} report(s): {len(all_changes)}")

    # build the Excel summary automatically — only for the reports we actually ran
    summary_report.build_report(only_reports=processed)
    print("  Report saved: output/Monthly_Quality_Check_Report.xlsx")
    print(f"{'='*55}")