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

def analyze_exclusive_type(series_str, total_rows):
    """
    Analyzes cell character structures purely blindly.
    Enforces a strict waterfall strategy to classify columns into exactly ONE type.
    """
    filled_count = len(series_str)
    if filled_count == 0:
        return "text", {}

    # 1. Date Check Waterfall (Regex character structure matching)
    date_regexes = [
        r'^\d{4}[-/]\d{2}[-/]\d{2}$',
        r'^\d{2}[-/]\d{2}[-/]\d{2,4}$'
    ]
    date_matches = series_str.apply(lambda x: any(re.match(r, x) for r in date_regexes)).sum()
    if (date_matches / max(1, filled_count)) > 0.4:
        unique_masks = series_str.apply(lambda x: re.sub(r'\d', 'X', x)).nunique()
        has_mixed = 1 if unique_masks > 1 else 0
        return "date", {"has_mixed_formats": has_mixed}

    # 2. Phone Check Waterfall (Digit spans subject to numeric zero truncation)
    cleaned_digits = series_str.apply(lambda x: re.sub(r'[\s\-\(\)\+]', '', x))
    phone_matches = cleaned_digits.apply(lambda x: x.isdigit() and (7 <= len(x) <= 15)).sum()
    if (phone_matches / max(1, filled_count)) > 0.5 and not series_str.str.contains('@').any():
        has_missing_zero = 1 if series_str.apply(lambda x: x.startswith(('1','2','3','4','5','6','7','8','9')) and not x.startswith('+')).any() else 0
        return "phone", {"has_missing_zeros": has_missing_zero}

    # 3. Number Check Waterfall (Handles floats, text numbers, currency markers, negative symbols)
    number_score = 0
    has_contamination = 0
    text_numbers = {'one', 'two', 'three', 'ten', 'twenty', 'thirty', 'forty', 'fifty'}
    
    for val in series_str:
        cleaned = re.sub(r'[^\d\.\-]', '', val)
        if re.match(r'^-?\d+(\.\d+)?$', cleaned):
            number_score += 1
            if re.search(r'[A-Za-z\$£€R]', val) or '-' in val:
                has_contamination = 1
        elif val.lower() in text_numbers:
            number_score += 1
            has_contamination = 1

    if (number_score / max(1, filled_count)) > 0.4 and not series_str.str.contains('@').any():
        return "number", {"has_contamination": has_contamination}

    # 4. Text Fallback Default
    lower_c = series_str.apply(lambda x: x.islower()).sum()
    upper_c = series_str.apply(lambda x: x.isupper()).sum()
    
    if upper_c > lower_c and upper_c > (filled_count * 0.5):
        smart_case = "uppercase"
    elif lower_c > upper_c and lower_c > (filled_count * 0.5):
        smart_case = "lowercase"
    else:
        smart_case = "titlecase"
        
    return "text", {"smart_case": smart_case}


def run_table_similarity_scan(df):
    """
    Scans the entire table matrix to group textual typo clusters (e.g. Unknow -> Unknown)
    """
    text_pool = []
    for col in df.columns:
        text_pool.extend(df[col].dropna().astype(str).str.strip().unique())
        
    unique_tokens = list(set(text_pool))
    typos = []
    
    for idx, word in enumerate(unique_tokens):
        for candidate in unique_tokens[idx+1:]:
            if len(word) > 4 and len(candidate) > 4:
                if word[:-1] == candidate or candidate[:-1] == word or (word.lower() != candidate.lower() and word.lower()[:5] == candidate.lower()[:5]):
                    typos.append({"flagged_value": word, "suggested_fix": candidate})
                    if len(typos) >= 8:
                        return typos
    return typos


@app.route("/parse-excel", methods=["POST"])
def parse_excel():
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file part in the request"}), 400

        file = request.files["file"]
        if file.filename == "":
            return jsonify({"error": "No selected file"}), 400

        # Force elements to string directly to capture leading formatting anomalies
        if file.filename.endswith('.csv'):
            df = pd.read_csv(io.BytesIO(file.read()), dtype=str)
        else:
            df = pd.read_excel(io.BytesIO(file.read()), dtype=str)

        # Map header columns exactly as before
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

            # Global Table Layout Shift Analyzer
            if total_rows > 5 and filled_count > 0 and (filled_count / total_rows) < 0.15:
                layout_shifts.append({
                    "column": col,
                    "error_msg": f"Stray text detected in {col}. Data may have shifted out of bounds.",
                    "sample_value": col_str.iloc[0] if len(col_str) > 0 else ""
                })

            blank_count = int(series.isna().sum() + (series.astype(str).str.strip() == "").sum())
            
            # Type evaluations
            detected_type, metrics = analyze_exclusive_type(col_str, total_rows)

            # Define smart type default tokens
            suggested_token = "Unknown"
            if detected_type == "number":
                suggested_token = "0"
            elif detected_type == "date":
                suggested_token = "None"

            # Enforce 5 Core Business Topics JSON Schema output
            column_diagnostics[col] = {
                "detected_primary_type": detected_type,
                
                "topic_1_empty_cells": {
                    "mistake_tracking": {
                        "empty_cell_count": blank_count,
                        "has_missing_values": 1 if blank_count > 0 else 0
                    },
                    "smart_defaults_and_options": {
                        "suggested_default_token": suggested_token
                    }
                },
                
                "topic_2_inconsistent_dates": {
                    "mistake_tracking": {
                        "is_applicable_date_column": 1 if detected_type == "date" else 0,
                        "has_mixed_formats": metrics.get("has_mixed_formats", 0)
                    },
                    "smart_defaults_and_options": {
                        "preselected_universal_fallback": "YYYY-MM-DD",
                        "regional_format_choices": ["YYYY-MM-DD", "DD-MM-YYYY", "MM-DD-YYYY", "YYYY/MM/DD"]
                    }
                },
                
                "topic_3_phone_numbers": {
                    "mistake_tracking": {
                        "is_applicable_phone_column": 1 if detected_type == "phone" else 0,
                        "has_truncated_leading_zeros": metrics.get("has_missing_zeros", 0)
                    },
                    "smart_defaults_and_options": {
                        "preselected_mode_configuration": "single",
                        "mode_choices": ["single", "mixed"]
                    }
                },
                
                "topic_4_text_cleaning": {
                    "subtopic_1_formatting": {
                        "is_applicable_text_column": 1 if detected_type == "text" else 0,
                        "prechecked_smart_case_option": metrics.get("smart_case", "titlecase"),
                        "casing_selector_choices": ["lowercase", "uppercase", "titlecase"]
                    },
                    "subtopic_2_typos": {
                        "global_table_scan_typos_found": global_typos,
                        "prechecked_master_toggle_clear_all": 1 if len(global_typos) > 0 else 0
                    }
                },
                
                "topic_5_number_cleaning": {
                    "mistake_tracking": {
                        "is_applicable_number_column": 1 if detected_type == "number" else 0,
                        "has_contamination_artifacts": metrics.get("has_contamination", 0)
                    },
                    "smart_defaults_and_options": {
                        "strip_formatting_preserve_signs": 1,
                        "preselected_rounding_uniformity_decimal": 2
                    }
                }
            }

        # Your exact pipe-delimited output configuration restored
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
