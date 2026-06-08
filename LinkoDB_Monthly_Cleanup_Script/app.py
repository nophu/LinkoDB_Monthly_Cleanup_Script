"""
app.py — a web version of the Linko Monthly Cleanup tool.


It reuses your existing pipeline (rubric_parser, data_parser, data_validator,
summary_report) — no logic is duplicated here, this file only adds the web UI.
"""
import os
import json
import tempfile

import pandas as pd
import streamlit as st

import rubric_parser
import data_parser
import data_validator
import summary_report
from config import REPORT_CONFIG, match_report


# core pipeline (no Streamlit code in here, so it's easy to test)
def run_pipeline(rubric_path, data_paths):
    """Run the full cleanup on saved file paths. Returns (changes, processed, skipped, report_path)."""
    os.makedirs("output", exist_ok=True)
    rubric = rubric_parser.parse_rubric(rubric_path)

    all_changes, processed, skipped = [], [], []
    for path in data_paths:
        filename = os.path.basename(path)
        key, cfg = match_report(filename)
        if cfg is None:
            skipped.append(filename)
            continue
        records = data_parser.parse_data(path, rubric)
        _, changes = data_validator.validate_data(
            records, rubric, key, only_fields=cfg["fields"]
        )
        all_changes.extend(changes)
        processed.append(key)

    with open("output/all_changes.json", "w") as f:    json.dump(all_changes, f, indent=2, default=str)

    report_path = "output/Monthly_Quality_Check_Report.xlsx"
    summary_report.build_report(only_reports=processed)
    return all_changes, processed, skipped, report_path


# web interface
st.set_page_config(page_title="Linko Monthly Cleanup", page_icon="🧹", layout="wide")

st.title("🧹 Linko DB Monthly Quality Check")
st.write(
    "Upload your rubric and the report file(s) you want to check. "
    "The tool validates each field against the rubric and builds a color-coded "
    "Excel report you can download."
)

col1, col2 = st.columns(2)
with col1:
    st.subheader("1. Rubric")
    rubric_file = st.file_uploader(
        "Final Scripts and Update Plan", type=["xlsx", "xls"], key="rubric"
    )
with col2:
    st.subheader("2. Report file(s)")
    data_files = st.file_uploader(
        "Pick one or more (Permit List, FSE, Master List, AG, Events)",
        type=["xlsx", "xls"], accept_multiple_files=True, key="data"
    )
ready = rubric_file is not None and bool(data_files)

if st.button("Run Quality Check", type="primary", disabled=not ready):
    with st.spinner("Checking your files..."):
        # write the uploads to a temp folder so the existing code can read them by path
        with tempfile.TemporaryDirectory() as tmp:
            rubric_path = os.path.join(tmp, rubric_file.name)
            with open(rubric_path, "wb") as f:  f.write(rubric_file.getbuffer())

            data_paths = []
            for uf in data_files:
                p = os.path.join(tmp, uf.name)
                with open(p, "wb") as f:  f.write(uf.getbuffer())
                data_paths.append(p)

            changes, processed, skipped, report_path = run_pipeline(rubric_path, data_paths)

    # results
    st.success(f"Done! Found {len(changes)} issue(s) across {len(processed)} report(s).")

    if skipped:
        st.warning(
            "These files didn't match a known report and were skipped: "
            + ", ".join(skipped)
        )

    # download button for the Excel report
    with open(report_path, "rb") as f:
        st.download_button(
            "⬇️  Download the Excel report",
            data=f,
            file_name="Monthly_Quality_Check_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

    # quick on-screen summary table
    if changes:
        df = pd.DataFrame(changes)
        df["result"] = df["status"].map({"fixed": "🟢 Auto-fixed", "flagged": "🟡 Needs review"})
        show = df[["facility", "permit_no", "field", "original", "result", "note"]].rename(
            columns={
                "facility": "Facility", "permit_no": "Permit",
                "field": "Field", "original": "Current value",
                "result": "Result", "note": "Details",
            }
        )

        fixed_n   = (df["status"] == "fixed").sum()
        flagged_n = (df["status"] == "flagged").sum()
        m1, m2, m3 = st.columns(3)
        m1.metric("Total issues", len(df))
        m2.metric("🟢 Auto-fixed", int(fixed_n))
        m3.metric("🟡 Needs review", int(flagged_n))
        st.dataframe(show, use_container_width=True, hide_index=True)
    else:  st.info("No issues found — everything is clean!")
else:
    if not ready:  st.caption("Upload a rubric and at least one report file to enable the button.")