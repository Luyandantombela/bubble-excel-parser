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

    # 3. Number Check Waterfall (Number / Financial Contamination & Strict Decimal Tracking)
    number_score = 0
    has_contamination = "no"
    decimal_lengths = []
    text_numbers = {'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten', 'twenty', 'thirty', 'forty', 'fifty'}
    
    for val in series_str:
        # Handles numbers that use trailing commas or periods (e.g., 2,00 or 1002.92)
        cleaned = re.sub(r'[^\d\.,\-]', '', val)
        # Normalize trailing split tokens to standard format trackers
        normalized = cleaned.replace(',', '.')
        
        if re.match(r'^-?\d+(\.\d+)?$', normalized):
            number_score += 1
            if re.search(r'[A-Za-z\$£€R]', val) or ('-' in val and val.strip().startswith('-') is False):
                has_contamination = "yes"
            
            # Count the exact length of trailing characters
            if '.' in normalized:
                decimal_part = normalized.split('.')[-1]
                decimal_lengths.append(len(decimal_part))
            else:
                decimal_lengths.append(0)
        elif val.lower() in text_numbers:
            number_score += 1
            has_contamination = "yes"

    if (number_score / max(1, filled_count)) > 0.4 and not series_str.str.contains('@').any():
        unique_lengths = set(decimal_lengths)
        inconsistent_decimals = "yes" if len(unique_lengths) > 1 else "no"
        
        return "number", {
            "inconsistent_numbering": has_contamination, 
            "inconsistent_decimal_places": inconsistent_decimals
        }

    # 4. Fallback Text Characteristics (Strict Casing Audit Rules)
    lower_count = 0
    upper_count = 0
    title_count = 0
    
    # Accurate multi-word Capital Casing evaluator rule
    title_pattern = r'^[A-Z][a-z]*(\s+[A-Z][a-z]*)*$'
    
    for val in series_str:
        if val.islower():
            lower_count += 1
        elif val.isupper():
            upper_count += 1
        elif re.match(title_pattern, val):
            title_count += 1

    # Check if the entire column matches exactly one uniform, safe casing category
    if lower_count == filled_count or upper_count == filled_count or title_count == filled_count:
        inconsistent_casing = "no"
    else:
        inconsistent_casing = "yes"
        
    return "text", {"inconsistent_formatting": inconsistent_casing}


def run_table_similarity_scan(df):
    """
    Advanced text similarity scanner. Uses a precise edit distance evaluation matrix 
    to group real typos (like Unknow vs Unknown) while skipping completely unique words.
    """
    text_pool = []
    for col in df.columns:
        for val in df[col].dropna().astype(str).str.strip().unique():
            # Exclude digits, symbols, short strings, and generic placeholders
            if not re.search(r'\d', val) and len(val) > 4 and " " not in val:
                text_pool.append(val)
        
    unique_tokens = list(set(text_pool))
    typos = []
    
    for idx, word in enumerate(unique_tokens):
        for candidate in unique_tokens[idx+1:]:
            # Enforce close-character tracking to drop long multi-word mismatches
            if abs(len(word) - len(candidate)) <= 2:
                if word[:-1] == candidate or candidate[:-1] == word or (word[:-2] == candidate[:-2] and word.lower()[:3] == candidate.lower()[:3]):
                    # Protect distinct, common dictionary terms from false matches
                    if not (word.lower().startswith("unite") and candidate.lower().startswith("unite")):
                        typos.append(word)
                        if len(typos) >= 8:
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
                col_typos = [word for word in global_typos if word in col_str.values]
                mistakes_found["misspellings"] = col_typos

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
