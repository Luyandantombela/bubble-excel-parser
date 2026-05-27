import io
import json
import os
import re
import pandas as pd
from flask import Flask, jsonify, request
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

def is_numeric_column(series):
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

        df.columns = [
            str(col).strip() if not str(col).startswith("Unnamed:") else f"Column {i+1}"
            for i, col in enumerate(df.columns)
        ]
        headers = list(df.columns)
        total_rows = len(df)

        column_diagnostics = {}
        layout_shifts = []

        for i, col in enumerate(headers):
            series = df[col]
            col_str = series.dropna().astype(str).str.strip()
            filled_rows_count = len(col_str)

            if filled_rows_count > 0 and (filled_rows_count / max(1, total_rows)) < 0.15:
                if i >= (len(headers) - 2):
                    layout_shifts.append({
                        "column": col,
                        "error_msg": f"Stray text detected in {col}. Data may have shifted out of bounds.",
                        "sample_value": col_str.iloc[0] if len(col_str) > 0 else ""
                    })

            blank_count = series.isna().sum() + (series.astype(str).str.strip() == "").sum()
            has_blanks = int(blank_count > 0)

            digits_only = col_str.str.replace(r'\D', '', regex=True)
            numeric_like_count = col_str.str.contains(r'[\d\-\.\$\(R]', regex=True).sum()
            
            is_num = int((numeric_like_count / max(1, filled_rows_count)) > 0.4 and not col_str.str.contains(r'[@]', regex=True).any())
            is_phone = int(digits_only.str.len().isin(range(3, 16)).sum() > 0.15 * max(1, filled_rows_count) and not is_num and "@" not in "".join(col_str))
            is_date = int(col_str.str.contains(r'(\d{4}[-/]\d{2}[-/]\d{2})|(\d{2}[-/]\d{2}[-/]\d{4})|(\d{2}[-/]\d{2}[-/]\d{2})', regex=True).sum() > 0.15 * max(1, filled_rows_count))
            is_text = int(not is_num and not is_date and not is_phone)

            default_fill = "0" if is_num else "Unknown"

            date_mismatch = 0
            if is_date:
                parsed_dates = pd.to_datetime(col_str, errors='coerce')
                date_mismatch = int(parsed_dates.isna().sum() > 0)

            phone_missing_zero = 0
            if is_phone:
                phone_missing_zero = int((digits_only.str.len() == 9).sum() > 0 or (digits_only.str.startswith(('7', '8', '6'))).sum() > 0)

            has_text_chaos = 0
            default_case_choice = "title"
            if is_text:
                has_mixed_case = int(col_str.str.isupper().sum() > 0 and col_str.str.islower().sum() > 0)
                has_spaces = int((col_str.str.startswith(" ") | col_str.str.endswith(" ")).sum() > 0)
                if has_mixed_case or has_spaces:
                    has_text_chaos = 1
                if (col_str.str.isupper().sum() / max(1, filled_rows_count)) > 0.5:
                    default_case_choice = "upper"

            has_number_chaos = 0
            has_decimals = 0
            if is_num:
                has_currency_symbols = int(col_str.str.contains(r'[R\$a-zA-Z]', regex=True).sum() > 0)
                try:
                    # Clean the data safely into pure numbers before running checks
                    numeric_floats = pd.to_numeric(col_str.str.replace(r'[R\$\s,]', '', regex=True), errors='coerce').dropna()
                    has_decimals = int((numeric_floats % 1 != 0).sum() > 0)
                    
                    # ⭐ FIX: Check for negative values safely using our clean numbers list
                    has_negatives = int((numeric_floats < 0).any())
                except:
                    has_negatives = 0
                
                if has_currency_symbols or has_decimals or has_negatives:
                    has_number_chaos = 1

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
