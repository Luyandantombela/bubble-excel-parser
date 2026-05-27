import io, json, os, re
import pandas as pd
import numpy as np
from flask import Flask, jsonify, request
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

def analyze_exclusive_type(series_str):
    filled_count = len(series_str)
    if filled_count == 0: return "text", {}

    date_regexes = [r'^\d{4}[-/]\d{2}[-/]\d{2}$', r'^\d{2}[-/]\d{2}[-/]\d{2,4}$']
    date_matches = series_str.apply(lambda x: any(re.match(r, x) for r in date_regexes)).sum()
    if (date_matches / max(1, filled_count)) > 0.4:
        unique_masks = series_str.apply(lambda x: re.sub(r'\d', 'X', x)).nunique()
        has_mixed = "yes" if unique_masks > 1 else "no"
        return "date", {"inconsistent_date_formatting": has_mixed, "desc": "Mixed date formats found." if has_mixed == "yes" else "Dates are uniform."}

    email_density = series_str.str.contains('@', regex=False).sum()
    if (email_density / max(1, filled_count)) > 0.4:
        has_invalid, has_mixed_case = "no", "no"
        invalid_list = []
        valid_email_regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$'
        for val in series_str:
            clean_val = val.strip()
            if not re.match(valid_email_regex, clean_val):
                has_invalid = "yes"
                if clean_val != "": invalid_list.append(clean_val)
            if any(c.isupper() for c in clean_val): has_mixed_case = "yes"
        return "email", {
            "invalid_emails": has_invalid, 
            "invalid_emails_desc": "Column contains broken email formats (e.g. missing a domain standard extension)." if has_invalid == "yes" else "Email formats are valid.", 
            "invalid_email_list": list(set(invalid_list)),
            "mixed_case_emails": has_mixed_case, 
            "mixed_case_emails_desc": "Emails contain mixed uppercase letters. These should be lowercase." if has_mixed_case == "yes" else "Email casing is uniform."
        }

    cleaned_digits = series_str.apply(lambda x: re.sub(r'[\s\-\(\)\+]', '', x))
    phone_matches = cleaned_digits.apply(lambda x: x.isdigit() and (7 <= len(x) <= 15)).sum()
    if (phone_matches / max(1, filled_count)) > 0.4:
        has_missing_zero = "yes" if series_str.apply(lambda x: x.startswith(('1','2','3','4','5','6','7','8','9')) and not x.startswith('+')).any() else "no"
        return "phone", {"missing_leading_zeros": has_missing_zero, "desc": "Phone numbers are truncated missing leading zeros." if has_missing_zero == "yes" else "Phone numbers are uniform."}

    number_score, has_contamination, decimal_lengths = 0, "no", []
    text_numbers = {'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten', 'twenty', 'thirty', 'forty', 'fifty'}
    for val in series_str:
        cleaned = re.sub(r'[^\d\.,\-]', '', val).replace(',', '.')
        if re.match(r'^-?\d+(\.\d+)?$', cleaned):
            number_score += 1
            if re.search(r'[A-Za-z\$£€R]', val) or ('-' in val and val.strip().startswith('-') is False): has_contamination = "yes"
            decimal_lengths.append(len(cleaned.split('.')[-1]) if '.' in cleaned else 0)
        elif val.lower() in text_numbers:
            number_score += 1; has_contamination = "yes"

    if (number_score / max(1, filled_count)) > 0.4:
        inconsistent_decimals = "yes" if len(set(decimal_lengths)) > 1 else "no"
        return "number", {"inconsistent_numbering": has_contamination, "numbering_desc": "Math cells contain currency symbols or text characters." if has_contamination == "yes" else "Numbers are cleanly formatted.", "inconsistent_decimal_places": inconsistent_decimals, "decimal_desc": "Uneven decimal lengths found across numbers." if inconsistent_decimals == "yes" else "Decimal lengths are uniform."}

    lower_count, upper_count, title_count, has_newlines = 0, 0, 0, "no"
    title_pattern = r'^[A-Z][a-z]*(\s+[A-Z][a-z]*)*$'
    for val in series_str:
        if "\n" in val or "\r" in val: has_newlines = "yes"
        if val.islower(): lower_count += 1
        elif val.isupper(): upper_count += 1
        elif re.match(title_pattern, val): title_count += 1

    if lower_count == filled_count or upper_count == filled_count or title_count == filled_count: inconsistent_casing = "no"
    else: inconsistent_casing = "yes"
    return "text", {"inconsistent_formatting": inconsistent_casing, "casing_desc": "Hidden newline breaks (\\n) found breaking cell format limits." if has_newlines == "yes" else "Inconsistent text casing layouts found."}

def run_table_similarity_scan(df):
    text_pool = []
    for col in df.columns:
        for val in df[col].dropna().astype(str).str.strip().unique():
            if not re.search(r'\d', val) and len(val) > 4 and " " not in val: text_pool.append(val)
    unique_tokens = list(set(text_pool))
    typos = []
    for idx, word in enumerate(unique_tokens):
        for candidate in unique_tokens[idx+1:]:
            if abs(len(word) - len(candidate)) <= 2:
                if word[:-1] == candidate or candidate[:-1] == word or (word[:-2] == candidate[:-2] and word.lower()[:3] == candidate.lower()[:3]):
                    if not (word.lower().startswith("unite") and candidate.lower().startswith("unite")):
                        typos.append(word)
                        if len(typos) >= 8: return typos
    return typos

@app.route("/parse-excel", methods=["POST"])
def parse_excel():
    try:
        if "file" not in request.files: return jsonify({"error": "No file chunk found"}), 400
        file = request.files["file"]
        if file.filename == "": return jsonify({"error": "Empty name"}), 400
        df = pd.read_csv(io.BytesIO(file.read()), dtype=str) if file.filename.endswith('.csv') else pd.read_excel(io.BytesIO(file.read()), dtype=str)
        df.columns = [str(col).strip() if not str(col).startswith("Unnamed:") else f"Column {i+1}" for i, col in enumerate(df.columns)]
        headers, total_rows, column_diagnostics, layout_shifts = list(df.columns), len(df), {}, []
        global_typos = run_table_similarity_scan(df)

        # 🛠️ GLOBAL DUPLICATE ROW SCANNER (Identifies exact string match duplicates)
        # Finds row positions that are identical copies (ignoring the first instance)
        duplicate_mask = df.duplicated(keep='first')
        # Map 0-based technical dataframe indexes directly into human 1-based spreadsheet rows
        duplicate_indices = [int(idx + 1) for idx, is_dup in enumerate(duplicate_mask) if is_dup]

        for i, col in enumerate(headers):
            series = df[col]
            col_str = series.dropna().astype(str).str.strip()
            filled_count = len(col_str)
            if total_rows > 5 and filled_count > 0 and (filled_count / total_rows) < 0.15:
                layout_shifts.append({"column": col, "error_msg": f"Stray text boundary shift in {col}.", "sample_value": col_str.iloc[0] if len(col_str) > 0 else ""})
            blank_count = int(series.isna().sum() + (series.astype(str).str.strip() == "").sum())
            detected_type, type_metrics = analyze_exclusive_type(col_str)
            mistakes_found = {"blank_cells": blank_count, "blank_cells_desc": f"Found {blank_count} empty rows missing values." if blank_count > 0 else "No missing values."}

            if detected_type == "number":
                mistakes_found["inconsistent_numbering"] = type_metrics.get("inconsistent_numbering", "no")
                mistakes_found["inconsistent_numbering_desc"] = type_metrics.get("numbering_desc", "")
                mistakes_found["inconsistent_decimal_places"] = type_metrics.get("inconsistent_decimal_places", "no")
                mistakes_found["inconsistent_decimal_places_desc"] = type_metrics.get("decimal_desc", "")
            elif detected_type == "date":
                mistakes_found["inconsistent_dates_formatting"] = type_metrics.get("inconsistent_date_formatting", "no")
                mistakes_found["inconsistent_dates_formatting_desc"] = type_metrics.get("desc", "")
            elif detected_type == "phone":
                mistakes_found["missing_leading_zeros"] = type_metrics.get("missing_leading_zeros", "no")
                mistakes_found["missing_leading_zeros_desc"] = type_metrics.get("desc", "")
            elif detected_type == "email":
                mistakes_found["invalid_emails"] = type_metrics.get("invalid_emails", "no")
                mistakes_found["invalid_emails_desc"] = type_metrics.get("invalid_emails_desc", "")
                mistakes_found["invalid_email_list"] = type_metrics.get("invalid_email_list", [])
                mistakes_found["mixed_case_emails"] = type_metrics.get("mixed_case_emails", "no")
                mistakes_found["mixed_case_emails_desc"] = type_metrics.get("mixed_case_emails_desc", "")
            elif detected_type == "text":
                mistakes_found["inconsistent_formatting"] = type_metrics.get("inconsistent_formatting", "no")
                mistakes_found["inconsistent_formatting_desc"] = type_metrics.get("casing_desc", "")
                mistakes_found["misspellings"] = [w for w in global_typos if w in col_str.values]
            column_diagnostics[col] = {"class": detected_type, "mistakes_found": mistakes_found}

        df_cleaned = df.fillna("")
        clean_rows = ["|".join([str(val) for val in row]) for _, row in df_cleaned.iterrows()]
        
        # 📦 Returns the duplicate rows inside your primary payload packet
        return jsonify({
            "headers": headers, 
            "rows_json": clean_rows, 
            "layout_alignment_errors": layout_shifts, 
            "duplicate_row_indices": duplicate_indices,
            "diagnostics": column_diagnostics
        }), 200
    except Exception as e:
        return jsonify({"error": f"Internal MasterX workflow crash: {str(e)}"}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
