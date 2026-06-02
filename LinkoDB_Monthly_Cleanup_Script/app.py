import os
import io
import json
import tempfile
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import pandas as pd

import rubric_parser
import data_parser
import data_validator

app = Flask(__name__)

# allow requests from GitHub Pages frontend
CORS(app)
# POST /validate
# Accepts two uploaded files: rubric and data
# Returns a JSON validation report
@app.route('/validate', methods=['POST'])
def validate():
    # check both files were uploaded
    if 'rubric' not in request.files or 'data' not in request.files: return jsonify({"error": "Please upload both a rubric file and a data file"}), 400

    rubric_file = request.files['rubric']
    data_file   = request.files['data']

    # save both files to a temporary folder
    with tempfile.TemporaryDirectory() as tmpdir:
        rubric_path = os.path.join(tmpdir, rubric_file.filename)
        data_path   = os.path.join(tmpdir, data_file.filename)

        rubric_file.save(rubric_path)
        data_file.save(data_path)

        # run the pipeline
        try:
            rubric    = rubric_parser.parse_rubric(rubric_path)
            records   = data_parser.parse_data(data_path, rubric)
            validated, changes = data_validator.validate_data( records, rubric, data_file.filename )
        except Exception as e: return jsonify({"error": str(e)}), 500

        # count results
        fixed   = [c for c in changes if c["status"] == "fixed"]
        flagged = [c for c in changes if c["status"] == "flagged"]

        # build the report
        report = {
            "total_records":   len(validated),
            "total_changes":   len(changes),
            "auto_fixed":      len(fixed),
            "flagged":         len(flagged),
            "fixed_details":   fixed,
            "flagged_details": flagged[:100],  # first 100 to keep response small
            "flagged_total":   len(flagged)
        }

        # save cleaned Excel to a temp file and encode for download
        cleaned_path = os.path.join(tmpdir, "cleaned_data.xlsx")
        df = pd.DataFrame(validated)
        df.to_excel(cleaned_path, index=False)

        # read the cleaned file into memory before tempdir is deleted
        with open(cleaned_path, "rb") as f:  cleaned_bytes = f.read()

    # return report as JSON, include cleaned file as base64
    import base64
    report["cleaned_file_b64"] = base64.b64encode(cleaned_bytes).decode("utf-8")
    report["cleaned_filename"] = f"cleaned_{data_file.filename}"
    return jsonify(report)

# GET /health
# Simple check to confirm the server is running
@app.route('/health', methods=['GET'])
def health():  return jsonify({"status": "ok"})

if __name__ == '__main__':  app.run(debug=True)