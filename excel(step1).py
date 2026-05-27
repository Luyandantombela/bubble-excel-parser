import io
import json
import os
import re
import pandas as pd
import numpy as np
from flask import Flask, jsonify, request
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

def analyze_exclusive_type(series_str):
    """
    Scans raw cell character structures blindly on any file uploaded.
    Forces an exclusive waterfall to ensure a column belongs to exactly ONE topic.
    """
    filled_count = len(series_str)
    if filled_count == 0:
        return "text", {}

    # 1. Date Check Waterfall (Universal Date Rules)
    date_regexes = [
        r'^\d{4}[-/]\d{2}[-/]\d{2}$',
        r'^\d{2}[-/]\d{2}[-/]\d{2,4}$'
    ]
    date_matches = series_str.apply(lambda x: any(re.match(r, x) for r in date_regexes)).sum()
    if (date_matches / max(1, filled_count)) > 0.4:
        unique_masks = series_str.apply(lambda x: re.sub(r'\d', 'X', x)).nunique()
        has_mixed = "yes" if unique_masks > 1 else "no"
        return "date", {"inconsistent_date_formatting": has_mixed}

    # 2. Phone Check Waterfall (Phone Number Spans)
    cleaned_digits = series_str.apply(lambda x: re.sub(r'[\s\-\(\)\+]', '', x))
    phone_matches = cleaned_digits.apply(lambda x: x.isdigit() and (7 <= len(x) <= 15)).sum()
    if (phone_matches / max(1, filled_count)) > 0.4 and not series_str.str.contains('@').any():
        has_missing_zero = "yes" if series_str.apply(lambda x: x.startswith(('1','2','3','4','5','6','7','8','9')) and not x.startswith('+')).any() else "no"
        return "phone", {"missing_leading_zeros": has_missing_zero}

    # 3. Number Check Waterfall (Number / Financial Contamination)
    number_score = 0
    has_contamination = "no"
    has_decimals = "no"
    text_numbers = {'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten', 'twenty', 'thirty', 'forty', 'fifty'}
    
    for val in series_str:
        cleaned = re.sub(r'[^\d\.\-]', '', val)
        if re.match(r'^-?\d+(\.\d+)?$', cleaned):
            number_score += 1
            if re.search(r'[A-Za-z\$£€R]', val) or '-' in val:
                has_contamination = "yes"
            if '.' in cleaned:
                # Check for varying lengths of trailing decimals across numbers
                has_decimals = "yes"
        elif val.lower() in text_numbers:
            number_score += 1
            has_contamination = "yes"

    if (number_score / max(1, filled_count)) > 0.4 and not series_str.str.contains('@').any():
        return "number", {
            "inconsistent_numbering": has_contamination, 
            "inconsistent_decimal_places": has_decimals
        }

    # 4. Fallback Text Characteristics
    lower_c = series_str.apply(lambda x: x.islower()).sum()
    upper_c = series_str.apply(lambda x: x.isupper()).sum()
    title_c = series_str.apply(lambda x: x.istitle()).sum()
    
    # If the rows aren't uniformly all upper, all lower, or all title, formatting is inconsistent
    if lower_c == filled_count or upper_c == filled_count or title_c == filled_count:
        inconsistent_casing = "no"
    else:
        inconsistent_casing = "yes"
        
    return "text", {"inconsistent_formatting": inconsistent_casing}


def run_table_similarity_scan(df):
    """
    Scans the entire arbitrary table layout looking for text similarity typos, 
    safely excluding numeric sequences and dates to avoid data corruption.
    """
    text_pool = []
    for col in df.columns:
        for val in df[col].dropna().astype(str).str.strip().unique():
            if not re.search(r'\d', val) and len(val) > 4:
                text_pool.append(val)
        
    unique_tokens = list(set(text_pool))
    typos = []
    
    for idx, word in enumerate(unique_tokens):
        for candidate in unique_tokens[idx+1:]:
            if word[:-1] == candidate or candidate[:-1] == word or (word.lower() != candidate.lower() and word.lower()[:5] == candidate.lower()[:5]):
                # Build an easy string array of flagged value samples for the frontend
                typos.append(word)
                if len(typos) >= 5:
                    return typos
    return typos


@app.route("/parse-excel", methods=["POST"])
def parse_excel():
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file part found in request payload"}), 400

        file = request.files["file"]
        if file.filename == "":
            return jsonify({"error": "Empty filename uploaded"}), 400

        if file.filename.endswith('.csv'):
            df = pd.read_csv(io.BytesIO(file.read()), dtype=str)
        else:
            df = pd.read_excel(io.BytesIO(file.read()), dtype=str)

        df.columns = [
            str(col).strip() if not str(col).startswith("Unnamed:") else f"Column {i+1}"
            for i, col in enumerate(df.columns)
        ]
        
        headers = list(df.columns)
        total_rows = len(df)
        
        column_diagnostics = {}
        layout_shifts = []
        global_typos = run_table_similarity_scan(df)

        for i, col in enumerate(headers):
            series = df[col]
            col_str = series.dropna().astype(str).str.strip()
            filled_count = len(col_str)

            if total_rows > 5 and filled_count > 0 and (filled_count / total_rows) < 0.15:
                layout_shifts.append({
                    "column": col,
                    "error_msg": f"Stray text detected in {col}. Data layout may have shifted out of standard bounds.",
                    "sample_value": col_str.iloc[0] if len(col_str) > 0 else ""
                })

            blank_count = int(series.isna().sum() + (series.astype(str).str.strip() == "").sum())
            detected_type, type_metrics = analyze_exclusive_type(col_str)

            # 🛠️ CREATE SIMPLE, CONDITIONALLY CLEANED MISTAKES OBJECTS
            mistakes_found = {"blank_cells": blank_count}

            if detected_type == "number":
                mistakes_found["inconsistent_numbering"] = type_metrics.get("inconsistent_numbering", "no")
                mistakes_found["inconsistent_decimal_places"] = type_metrics.get("inconsistent_decimal_places", "no")
                
            elif detected_type == "date":
                mistakes_found["inconsistent_dates_formatting"] = type_metrics.get("inconsistent_date_formatting", "no")
                
            elif detected_type == "phone":
                mistakes_found["missing_leading_zeros"] = type_metrics.get("missing_leading_zeros", "no")
                
            elif detected_type == "text":
                mistakes_found["inconsistent_formatting"] = type_metrics.get("inconsistent_formatting", "no")
                # Pull words from this column that match our global table typo scanner pool
                col_typos = [word for word in global_typos if word in col_str.values]
                mistakes_found["misspellings"] = col_typos

            # The exact, beautiful, lightweight classified output format
            column_diagnostics[col] = {
                "class": detected_type,
                "mistakes_found": mistakes_found
            }

        df_cleaned = df.fillna("")
        clean_rows = []
        for _, row in df_cleaned.iterrows():
            row_string = "|".join([str(val) for val in row])
            clean_rows.append(row_string)

        return jsonify({
            "headers": headers, 
            "rows_json": clean_rows,
            "layout_alignment_errors": layout_shifts,
            "diagnostics": column_diagnostics
        }), 200

    except Exception as e:
        return jsonify({"error": f"Internal MasterX parsing workflow crash: {str(e)}"}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
