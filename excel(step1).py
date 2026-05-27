import io
import json
import os
import re
import pandas as pd
from flask import Flask, jsonify, request
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

@app.route("/parse-excel", methods=["POST"])
def parse_excel():
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file part in the request"}), 400

        file = request.files["file"]
        if file.filename == "":
            return jsonify({"error": "No selected file"}), 400

        file_bytes = file.read()
        file_stream = io.BytesIO(file_bytes)

        try:
            df = pd.read_excel(file_stream)
        except Exception:
            file_stream.seek(0)
            df = pd.read_excel(file_stream, engine="openpyxl")

        # Standardise headers cleanly
        df.columns = [
            str(col).strip() if not str(col).startswith("Unnamed:") else f"Column {i+1}"
            for i, col in enumerate(df.columns)
        ]
        headers = list(df.columns)
        total_rows = len(df)

        # -------------------------------------------------------------------
        # 🧠 THE BLIND DATA DIAGNOSTIC ENGINE (Analyzes cells, ignores names)
        # -------------------------------------------------------------------
        column_diagnostics = {}
        layout_shifts = []

        for i, col in enumerate(headers):
            series = df[col]
            col_str = series.dropna().astype(str).str.strip()
            filled_rows_count = len(col_str)

            # --- BUG FIX 3: GLOBAL LAYOUT SHIFT DETECTOR ---
            # If a column is almost entirely empty but contains a rare rogue entry out on the edge
            if filled_rows_count > 0 and (filled_rows_count / max(1, total_rows)) < 0.15:
                # If it's one of the last columns, flag it as a layout alignment error
                if i >= (len(headers) - 2):
                    layout_shifts.append({
                        "column": col,
                        "error_msg": f"Stray text detected in {col}. Data may have shifted out of bounds.",
                        "sample_value": col_str.iloc[0] if len(col_str) > 0 else ""
                    })

            # 1. EMPTY CELLS SCANNER (Applies to all columns universaly)
            blank_count = series.isna().sum() + (series.astype(str).str.strip() == "").sum()
            has_blanks = int(blank_count > 0)

            # --- BLIND DATA CLASSIFICATION SYSTEM ---
            # Look at characters inside the cells to figure out the column type natively
            digits_only = col_str.str.replace(r'\D', '', regex=True)
            has_digits_count = (digits_only.str.len() > 0).sum()
            
            # Check for math symbols or digits to classify Numeric Columns
            numeric_like_count = col_str.str.contains(r'[\d\-\.\$\(R]', regex=True).sum()
            
            is_num = int((numeric_like_count / max(1, filled_rows_count)) > 0.4 and not col_str.str.contains(r'[@]', regex=True).any())
            
            # --- BUG FIX 2: SENSITIVE PHONE DETECTOR ---
            # Checks if cells hold strings containing between 3 and 15 digits (handles short codes & long text noise)
            is_phone = int(digits_only.str.len().isin(range(3, 16)).sum() > 0.15 * max(1, filled_rows_count) and not is_num and "@" not in "".join(col_str))

            # Check for date patterns matching common dash/slash configurations
            is_date = int(col_str.str.contains(r'(\d{4}[-/]\d{2}[-/]\d{2})|(\d{2}[-/]\d{2}[-/]\d{4})|(\d{2}[-/]\d{2}[-/]\d{2})', regex=True).sum() > 0.15 * max(1, filled_rows_count))
            
            is_text = int(not is_num and not is_date and not is_phone)

            # Smart Default Filler Input Value Choice Assignment
            default_fill = "0" if is_num else "Unknown"

            # 2. DATE CONFORMANCE SCANNER
            date_mismatch = 0
            if is_date:
                parsed_dates = pd.to_datetime(col_str, errors='coerce')
                date_mismatch = int(parsed_dates.isna().sum() > 0)

            # Phone missing zeros check
            phone_missing_zero = 0
            if is_phone:
                # If it's a 9 digit string or starts with standard mobile digits without a 0
                phone_missing_zero = int((digits_only.str.len() == 9).sum() > 0 or (digits_only.str.startswith(('7', '8', '6'))).sum() > 0)

            # 4. TEXT FORMATTING SCANNER
            has_text_chaos = 0
            default_case_choice = "title"
            if is_text:
                has_mixed_case = int(col_str.str.isupper().sum() > 0 and col_str.str.islower().sum() > 0)
                has_spaces = int((col_str.str.startswith(" ") | col_str.str.endswith(" ")).sum() > 0)
                if has_mixed_case or has_spaces:
                    has_text_chaos = 1
                if (col_str.str.isupper().sum() / max(1, filled_rows_count)) > 0.5:
                    default_case_choice = "upper"

            # --- BUG FIX 1: UNIVERSAL COERCED NUMBER SCANNER ---
            # Strips typical characters dynamically to catch currency prefixes and decimal spacing
            has_number_chaos = 0
            has_decimals = 0
            if is_num:
                has_currency_symbols = int(col_str.str.contains(r'[R\$a-zA-Z]', regex=True).sum() > 0)
                try:
                    numeric_floats = pd.to_numeric(col_str.str.replace(r'[R\$\s,]', '', regex=True), errors='coerce').dropna()
                    has_decimals = int((numeric_floats % 1 != 0).sum() > 0)
                except: pass
                
                if has_currency_symbols or has_decimals or (series < 0).any():
                    has_number_chaos = 1

            # Compile the specific blind profile diagnostics dictionary entry
            column_diagnostics[col] = {
                "empty_cells": {
                    "found": int(has_blanks),
                    "count": int(blank_count),
                    "default_value": default_fill
                },
                "dates": {
                    "is_date_column": int(is_date),
                    "has_mixed_formats": int(date_mismatch),
                    "default_format": "YYYY-MM-DD"
                },
                "phones": {
                    "is_phone_column": int(is_phone),
                    "has_missing_zeros": int(phone_missing_zero),
                    "default_mode": "single"
                },
                "text": {
                    "is_text_column": int(is_text),
                    "has_formatting_issues": int(has_text_chaos),
                    "default_case": default_case_choice
                },
                "numbers": {
                    "is_number_column": int(is_num),
                    "has_formatting_issues": int(has_number_chaos),
                    "default_round_decimals": 2 if has_decimals else 0
                }
            }

        # Packaging cell data payload
        df_cleaned = df.fillna("")
        clean_rows = []
        for _, row in df_cleaned.iterrows():
            row_string = "|".join([str(val) for val in row])
            clean_rows.append(row_string)

        return jsonify({
            "headers": headers, 
            "rows_json": clean_rows,
            "diagnostics": column_diagnostics,
            "layout_alignment_errors": layout_shifts
        }), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
