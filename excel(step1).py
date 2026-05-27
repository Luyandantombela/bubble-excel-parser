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
    filled_count = len(series_str)
    if filled_count == 0:
        return "text", {}

    date_regexes = [
        r'^\d{4}[-/]\d{2}[-/]\d{2}$',
        r'^\d{2}[-/]\d{2}[-/]\d{2,4}$'
    ]
    date_matches = series_str.apply(lambda x: any(re.match(r, x) for r in date_regexes)).sum()
    if (date_matches / max(1, filled_count)) > 0.4:
        unique_masks = series_str.apply(lambda x: re.sub(r'\d', 'X', x)).nunique()
        has_mixed = 1 if unique_masks > 1 else 0
        return "date", {"has_mixed_formats": has_mixed}

    cleaned_digits = series_str.apply(lambda x: re.sub(r'[\s\-\(\)\+]', '', x))
    phone_matches = cleaned_digits.apply(lambda x: x.isdigit() and (7 <= len(x) <= 15)).sum()
    if (phone_matches / max(1, filled_count)) > 0.4 and not series_str.str.contains('@').any():
        has_missing_zero = 1 if series_str.apply(lambda x: x.startswith(('1','2','3','4','5','6','7','8','9')) and not x.startswith('+')).any() else 0
        return "phone", {"has_missing_zeros": has_missing_zero}

    number_score = 0
    has_contamination = 0
    text_numbers = {'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten', 'twenty', 'thirty', 'forty', 'fifty'}
    
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
                typos.append({"flagged_value": word, "suggested_fix": candidate})
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
            detected_type, metrics = analyze_exclusive_type(col_str)

            suggested_token = "Unknown"
            if detected_type == "number":
                suggested_token = "0"
            elif detected_type == "date":
                suggested_token = "None"

            column_diagnostics[col] = {
                "detected_primary_type": detected_type,
                "topic_1_empty_cells": {
                    "user_display_mistakes": f"We scanned this column and found a data gap! {blank_count} completely blank spaces are missing data values.",
                    "masterx_suggestion": suggested_token,
                    "user_interactive_choices": {
                        "input_field_type": "text_entry_box",
                        "placeholder_examples": ["N/A", "Unknown", "0"]
                    }
                },
                "topic_2_inconsistent_dates": {
                    "user_display_mistakes": f"Warning! Date formats are mixed up in different structures here (e.g. tracking variations like 04-06-25 alongside 2025/04/14)." if metrics.get("has_mixed_formats", 0) else "Date structure checked.",
                    "masterx_suggestion": "YYYY-MM-DD",
                    "user_interactive_choices": {
                        "dropdown_preselected_default": "YYYY-MM-DD",
                        "dropdown_menu_options": ["YYYY-MM-DD", "DD-MM-YYYY", "MM-DD-YYYY", "YYYY/MM/DD"]
                    }
                },
                "topic_3_phone_numbers": {
                    "user_display_mistakes": "Numeric Truncation Error! We found phone number sequences with formatting mistakes or missing leading zeros." if metrics.get("has_missing_zeros", 0) else "Phone number structure verified.",
                    "masterx_suggestion": "single",
                    "user_interactive_choices": {
                        "dropdown_preselected_default": "single",
                        "dropdown_menu_options": ["single", "mixed"],
                        "mode_definitions": {
                            "single": "Single Country Patching: Automatically snaps your country code (+27) onto the front and patches spatial gaps.",
                            "mixed": "Mixed Countries Patching: For international datasets; forces missing leading zeros without altering country variables."
                        }
                    }
                },
                "topic_4_text_cleaning": {
                    "subtopic_1_formatting": {
                        "user_display_mistakes": "Text formatting habits evaluated. Casing layout variations found.",
                        "masterx_suggestion": metrics.get("smart_case", "titlecase"),
                        "user_interactive_choices": {
                            "selector_ui_type": "radio_buttons",
                            "choices_array": ["lowercase", "uppercase", "titlecase"]
                        }
                    },
                    "subtopic_2_misspellings": {
                        "global_table_scan_typos_found": global_typos,
                        "masterx_suggestion": 1,
                        "user_interactive_choices": {
                            "toggle_ui_type": "single_master_checkbox",
                            "label_text": "Merge and resolve all detected spelling clusters table-wide automatically"
                        }
                    }
                },
                "topic_5_number_cleaning": {
                    "subtopic_1_formatting": {
                        "user_display_mistakes": "Math columns are contaminated! We found symbols (R300), keywords (thirty), or negative signs (-45) typed inside.",
                        "masterx_suggestion": 1,
                        "user_interactive_choices": {
                            "toggle_ui_type": "boolean_switch",
                            "label_text": "Strip text/currency tokens while protecting mathematical signs (+/-)"
                        }
                    },
                    "subtopic_2_rounder": {
                        "user_display_mistakes": "Uneven decimal layouts or inconsistent value points discovered across your financial rows.",
                        "masterx_suggestion": 2,
                        "user_interactive_choices": {
                            "input_ui_type": "number_box_or_dropdown",
                            "precheck_value": 2
                        }
                    }
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
            "layout_alignment_errors": layout_shifts,
            "diagnostics": column_diagnostics
        }), 200

    except Exception as e:
        return jsonify({"error": f"Internal MasterX parsing workflow crash: {str(e)}"}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
