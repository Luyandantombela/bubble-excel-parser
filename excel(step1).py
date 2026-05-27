import io
import json
import os
import re
import pandas as pd
from flask import Flask, jsonify, request
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# Helper function to detect number column types
def is_numeric_column(series):
    # Try converting non-null values to numeric to see if it's a number column
    dropped = series.dropna()
    if len(dropped) == 0:
        return False
    converted = pd.to_numeric(dropped.astype(str).str.replace(r'[R\$\s,]', '', regex=True), errors='coerce')
    return converted.notna().sum() / len(dropped) > 0.5

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

        # Clean and standardise column headers
        df.columns = [
            str(col).strip() if not str(col).startswith("Unnamed:") else f"Column {i+1}"
            for i, col in enumerate(df.columns)
        ]
        headers = list(df.columns)

        # -------------------------------------------------------------------
        # 🧠 THE DIAGNOSTIC DISCOVERY ENGINE (Scans mistakes & makes choices)
        # -------------------------------------------------------------------
        column_diagnostics = {}

        for col in headers:
            series = df[col]
            total_rows = len(series)
            
            # 1. EMPTY CELLS SCANNER
            blank_count = series.isna().sum() + (series.astype(str).str.strip() == "").sum()
            has_blanks = int(blank_count > 0)
            # Professional Default Choice: 0 for numbers, "Unknown" for text strings
            is_num = is_numeric_column(series)
            default_fill = "0" if is_num else "Unknown"

            # 2. DATE CONFORMANCE SCANNER
            has_dates = 0
            date_mismatch = 0
            col_str = series.dropna().astype(str).str.strip()
            # Basic regex to spot common forward slash or dash date patterns
            date_patterns = col_str.str.contains(r'(\d{4}[-/]\d{2}[-/]\d{2})|(\d{2}[-/]\d{2}[-/]\d{4})|(\d{2}[-/]\d{2}[-/]\d{2})', regex=True)
            if date_patterns.sum() > 0.2 * total_rows: # If over 20% looks like dates, treat as date column
                has_dates = 1
                # Try parsing to see how many fail standard ISO structure natively
                parsed_dates = pd.to_datetime(col_str, errors='coerce')
                date_mismatch = int(parsed_dates.isna().sum() > 0)

            # 3. PHONE NUMBER INTEGRITY SCANNER
            has_phones = 0
            phone_missing_zero = 0
            # Strip spaces/symbols to check digit length distributions
            digits_only = col_str.str.replace(r'\D', '', regex=True)
            valid_len_check = digits_only.str.len().isin([9, 10])
            if valid_len_check.sum() > 0.3 * total_rows: # Treat as a contact phone number column
                has_phones = 1
                # Spot numbers starting with 7, 8, 6 or missing standard country/area leading tags
                missing_zero_pattern = digits_only.str.len() == 9
                phone_missing_zero = int(missing_zero_pattern.sum() > 0)

            # 4. TEXT CASE SETTINGS SCANNER
            has_text_chaos = 0
            default_case_choice = "title" # Default: Capitalise Each First Letter
            if not is_num and not has_dates and not has_phones:
                # Check for inconsistent formatting variations inside text
                has_mixed_case = int(series.dropna().astype(str).str.isupper().sum() > 0 and series.dropna().astype(str).str.islower().sum() > 0)
                has_spaces = int((series.dropna().astype(str).str.startswith(" ") | series.dropna().astype(str).str.endswith(" ")).sum() > 0)
                if has_mixed_case or has_spaces:
                    has_text_chaos = 1
                
                # Smart choice selector: If the column text looks like an upper-case code block, preserve it
                if series.dropna().astype(str).str.isupper().sum() / max(1, len(series.dropna())) > 0.6:
                    default_case_choice = "upper"

            # 5. NUMBER CURRENCY & ROUNDING SCANNER
            has_number_chaos = 0
            has_decimals = 0
            if is_num and not has_dates:
                # Check for text symbols mixed into currency like 'R' or '$'
                has_currency_symbols = int(series.dropna().astype(str).str.contains(r'[R\$a-zA-Z]', regex=True).sum() > 0)
                # Check if there are long uneven decimal strings that require rounding
                try:
                    numeric_floats = pd.to_numeric(df[col].astype(str).str.replace(r'[R\$\s,]', '', regex=True), errors='coerce').dropna()
                    has_decimals = int((numeric_floats % 1 != 0).sum() > 0)
                except: pass
                
                if has_currency_symbols or has_decimals:
                    has_number_chaos = 1

            # Store the full column diagnosis report back for Bubble
            column_diagnostics[col] = {
                "empty_cells": {
                    "found": int(has_blanks),
                    "count": int(blank_count),
                    "default_value": default_fill
                },
                "dates": {
                    "is_date_column": int(has_dates),
                    "has_mixed_formats": int(date_mismatch),
                    "default_format": "YYYY-MM-DD"
                },
                "phones": {
                    "is_phone_column": int(has_phones),
                    "has_missing_zeros": int(phone_missing_zero),
                    "default_mode": "single"
                },
                "text": {
                    "is_text_column": int(not is_num and not has_dates and not has_phones),
                    "has_formatting_issues": int(has_text_chaos),
                    "default_case": default_case_choice
                },
                "numbers": {
                    "is_number_column": int(is_num),
                    "has_formatting_issues": int(has_number_chaos),
                    "default_round_decimals": 2 if has_decimals else 0
                }
            }

        # -------------------------------------------------------------------
        # 📦 PACKAGING THE CELL STRINGS GRID (Your original pipe serializer)
        # -------------------------------------------------------------------
        df_cleaned = df.fillna("")
        clean_rows = []
        for _, row in df_cleaned.iterrows():
            row_string = "|".join([str(val) for val in row])
            clean_rows.append(row_string)

        # Send everything back inside a structured JSON response bundle
        return jsonify({
            "headers": headers, 
            "rows_json": clean_rows,
            "diagnostics": column_diagnostics
        }), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
