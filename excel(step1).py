import io, json, os, re
import pandas as pd
import numpy as np
from flask import Flask, jsonify, request
from flask_cors import CORS

try:
    from spellchecker import SpellChecker
    SPELLCHECK_AVAILABLE = True
except ImportError:
    SPELLCHECK_AVAILABLE = False

app = Flask(__name__)
CORS(app)

CUSTOM_WHITELIST = {
    'joburg', 'pretoria', 'durban', 'bongi', 'tina', 'chris', 'linda',
    'sarah', 'mike', 'amy', 'john', 'todo', 'asdf', 'idk'
}


# ─────────────────────────────────────────────
# SPELL CHECK  →  returns list of
# {"original": "wrord", "suggestion": "word"}
# ─────────────────────────────────────────────
def find_column_typos(series_str):
    """
    Returns a list of dicts: {original, suggestion}.
    If pyspellchecker is available we get real corrections;
    otherwise falls back to frequency-based heuristic (no suggestion).
    """
    if SPELLCHECK_AVAILABLE:
        spell = SpellChecker()
        found = {}          # original → best suggestion
        for val in series_str.dropna().astype(str).str.strip():
            clean_text = re.sub(r'[^a-zA-Z\s]', '', val)
            words = [w.lower() for w in clean_text.split() if len(w) > 2]
            unknown = spell.unknown(words)
            for word in unknown:
                if word not in CUSTOM_WHITELIST and word not in found:
                    correction = spell.correction(word)
                    found[word] = correction if correction and correction != word else None
        return [{"original": k, "suggestion": v} for k, v in found.items()]

    # Fallback: frequency-based heuristic (no correction available)
    word_pool = []
    for val in series_str.dropna().astype(str).str.strip():
        clean_text = re.sub(r'[^a-zA-Z\s]', '', val)
        for chunk in clean_text.split():
            if len(chunk) > 3:
                word_pool.append(chunk.lower())
    if not word_pool:
        return []
    freq_map = pd.Series(word_pool).value_counts().to_dict()
    unique_words = list(freq_map.keys())
    flagged = []
    for idx, word in enumerate(unique_words):
        for candidate in unique_words[idx + 1:]:
            if abs(len(word) - len(candidate)) <= 2:
                if (word[:-1] == candidate or candidate[:-1] == word or
                        (word[:-2] == candidate[:-2] and word[:3] == candidate[:3])):
                    if not (word.startswith("unite") or candidate.startswith("unite")):
                        rare = word if freq_map[word] < freq_map[candidate] else candidate
                        flagged.append(rare)
    return [{"original": w, "suggestion": None} for w in set(flagged)]


# ─────────────────────────────────────────────
# TYPE DETECTION  (unchanged logic, same return)
# ─────────────────────────────────────────────
def analyze_exclusive_type(series_str):
    filled_count = len(series_str)
    if filled_count == 0:
        return "text", {}

    # DATE
    date_regexes = [r'^\d{4}[-/]\d{2}[-/]\d{2}$', r'^\d{2}[-/]\d{2}[-/]\d{2,4}$']
    date_matches = series_str.apply(lambda x: any(re.match(r, x) for r in date_regexes)).sum()
    if (date_matches / max(1, filled_count)) > 0.4:
        unique_masks = series_str.apply(lambda x: re.sub(r'\d', 'X', x)).nunique()
        has_mixed = "yes" if unique_masks > 1 else "no"
        return "date", {
            "inconsistent_date_formatting": has_mixed,
            "desc": "Mixed date formats found." if has_mixed == "yes" else "Dates are uniform."
        }

    # EMAIL
    email_density = series_str.str.contains('@', regex=False).sum()
    if (email_density / max(1, filled_count)) > 0.4:
        has_invalid, has_mixed_case = "no", "no"
        invalid_list = []
        valid_email_regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$'
        for val in series_str:
            clean_val = val.strip()
            if not re.match(valid_email_regex, clean_val):
                has_invalid = "yes"
                if clean_val != "":
                    invalid_list.append(clean_val)
            if any(c.isupper() for c in clean_val):
                has_mixed_case = "yes"
        return "email", {
            "invalid_emails": has_invalid,
            "invalid_emails_desc": "Column contains broken email formats.",
            "invalid_email_list": list(set(invalid_list)),
            "mixed_case_emails": has_mixed_case,
            "mixed_case_emails_desc": "Emails contain mixed uppercase letters." if has_mixed_case == "yes" else "Email casing is uniform."
        }

    # PHONE
    cleaned_digits = series_str.apply(lambda x: re.sub(r'[\s\-\(\)\+]', '', x))
    phone_matches = cleaned_digits.apply(lambda x: x.isdigit() and (7 <= len(x) <= 15)).sum()
    if (phone_matches / max(1, filled_count)) > 0.4:
        has_issue, issue_desc = "no", "Phone numbers are uniform."
        for val in series_str:
            clean_val = val.strip()
            digits_only = re.sub(r'\D', '', clean_val)
            if '?' in clean_val or clean_val.isalpha():
                has_issue, issue_desc = "yes", "Phone numbers contain invalid placeholder text symbols (like '??')."
                break
            elif clean_val.startswith(('1','2','3','4','5','6','7','8','9')) and not clean_val.startswith('+'):
                has_issue, issue_desc = "yes", "Phone numbers are truncated, missing leading zeros."
                break
            elif len(digits_only) > 0 and len(digits_only) < 9:
                has_issue, issue_desc = "yes", "Phone numbers contain broken, short sequences missing digits."
                break
        return "phone", {"missing_leading_zeros": has_issue, "desc": issue_desc}

    # NUMBER
    number_score, has_contamination, decimal_lengths = 0, "no", []
    text_numbers = {'one','two','three','four','five','six','seven','eight','nine','ten','twenty','thirty','forty','fifty'}
    for val in series_str:
        cleaned = re.sub(r'[^\d\.,\-]', '', val).replace(',', '.')
        if re.match(r'^-?\d+(\.\d+)?$', cleaned):
            number_score += 1
            if re.search(r'[A-Za-z\$£€R]', val) or ('-' in val and val.strip().startswith('-') is False):
                has_contamination = "yes"
            decimal_lengths.append(len(cleaned.split('.')[-1]) if '.' in cleaned else 0)
        elif val.lower() in text_numbers:
            number_score += 1
            has_contamination = "yes"

    if (number_score / max(1, filled_count)) > 0.4:
        inconsistent_decimals = "yes" if len(set(decimal_lengths)) > 1 else "no"
        return "number", {
            "inconsistent_numbering": has_contamination,
            "numbering_desc": "Math cells contain currency symbols or text characters." if has_contamination == "yes" else "Numbers are cleanly formatted.",
            "inconsistent_decimal_places": inconsistent_decimals,
            "decimal_desc": "Uneven decimal lengths found across numbers." if inconsistent_decimals == "yes" else "Decimal lengths are uniform."
        }

    # TEXT
    lower_count, upper_count, title_count, has_newlines = 0, 0, 0, "no"
    title_pattern = r'^[A-Z][a-z]*(\s+[A-Z][a-z]*)*$'
    for val in series_str:
        if "\n" in val or "\r" in val:
            has_newlines = "yes"
        if val.islower():
            lower_count += 1
        elif val.isupper():
            upper_count += 1
        elif re.match(title_pattern, val):
            title_count += 1

    inconsistent_casing = "no" if (lower_count == filled_count or upper_count == filled_count or title_count == filled_count) else "yes"
    return "text", {
        "inconsistent_formatting": inconsistent_casing,
        "casing_desc": "Hidden newline breaks found breaking cell format limits." if has_newlines == "yes" else "Inconsistent text casing layouts found."
    }


# ─────────────────────────────────────────────
# NEW: build the "suggested clean" cell value
# ─────────────────────────────────────────────
def suggest_clean_value(val, detected_type, type_metrics, spell_map):
    """
    Returns a suggested cleaned string for a single cell value.
    spell_map: dict {original_lower → suggestion} for this column.
    """
    if val is None or str(val).strip() == "":
        return ""

    s = str(val).strip()

    if detected_type == "text":
        # Apply spelling corrections
        words = s.split()
        corrected = []
        for w in words:
            key = re.sub(r'[^a-zA-Z]', '', w).lower()
            if key in spell_map and spell_map[key]:
                # Preserve original casing pattern
                suggestion = spell_map[key]
                if w.isupper():
                    suggestion = suggestion.upper()
                elif w.istitle():
                    suggestion = suggestion.title()
                corrected.append(suggestion)
            else:
                corrected.append(w)
        s = " ".join(corrected)

        # Apply casing suggestion: default to Title Case for text
        if type_metrics.get("inconsistent_formatting") == "yes":
            s = s.title()

        return s

    if detected_type == "number":
        # Strip currency symbols and letters, normalise decimals to 2dp
        cleaned = re.sub(r'[^\d\.\-]', '', s.replace(',', '.'))
        try:
            num = float(cleaned)
            return f"{num:.2f}"
        except ValueError:
            return s

    if detected_type == "date":
        # Attempt to parse and reformat to YYYY/MM/DD
        for fmt in ('%Y/%m/%d', '%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y',
                    '%m-%d-%y', '%m/%d/%Y', '%d %b %Y', '%b %d, %Y',
                    '%m-%d-%Y'):
            try:
                import datetime
                d = datetime.datetime.strptime(s, fmt)
                return d.strftime('%Y/%m/%d')
            except ValueError:
                continue
        return s

    if detected_type == "email":
        return s.lower()

    if detected_type == "phone":
        digits = re.sub(r'\D', '', s)
        if digits and not s.startswith('+') and len(digits) >= 9:
            return '0' + digits if not digits.startswith('0') else digits
        return s

    return s


# ─────────────────────────────────────────────
# MAIN ROUTE
# ─────────────────────────────────────────────
@app.route("/parse-excel", methods=["POST"])
def parse_excel():
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file chunk found"}), 400
        file = request.files["file"]
        if file.filename == "":
            return jsonify({"error": "Empty name"}), 400

        fname = file.filename.lower()
        raw = file.read()

        # Support CSV, XLSX, XLS, ODS, TSV
        if fname.endswith('.csv'):
            df = pd.read_csv(io.BytesIO(raw), dtype=str)
        elif fname.endswith('.tsv'):
            df = pd.read_csv(io.BytesIO(raw), sep='\t', dtype=str)
        elif fname.endswith('.ods'):
            df = pd.read_excel(io.BytesIO(raw), dtype=str, engine='odf')
        else:
            df = pd.read_excel(io.BytesIO(raw), dtype=str)

        df.columns = [
            str(col).strip() if not str(col).startswith("Unnamed:") else f"Column {i+1}"
            for i, col in enumerate(df.columns)
        ]

        headers     = list(df.columns)
        total_rows  = len(df)
        column_diagnostics = {}
        layout_shifts      = []

        duplicate_mask    = df.duplicated(keep='first')
        duplicate_indices = [int(idx + 1) for idx, is_dup in enumerate(duplicate_mask) if is_dup]

        # Per-column spell maps so suggest_clean_value can use them
        spell_maps = {}

        for i, col in enumerate(headers):
            series    = df[col]
            col_str   = series.dropna().astype(str).str.strip()
            filled_count = len(col_str)

            if total_rows > 5 and filled_count > 0 and (filled_count / total_rows) < 0.15:
                layout_shifts.append({
                    "column": col,
                    "error_msg": f"Stray text boundary shift in {col}.",
                    "sample_value": col_str.iloc[0] if len(col_str) > 0 else ""
                })

            blank_count = int(series.isna().sum() + (series.astype(str).str.strip() == "").sum())
            detected_type, type_metrics = analyze_exclusive_type(col_str)

            # Spellcheck (text columns only)
            typos = []
            spell_map = {}
            if detected_type == "text":
                typos = find_column_typos(col_str)
                spell_map = {t["original"]: t["suggestion"] for t in typos}
            spell_maps[col] = spell_map

            mistakes_found = {
                "blank_cells": blank_count,
                "blank_cells_desc": f"Found {blank_count} empty rows missing values." if blank_count > 0 else "No missing values."
            }

            if detected_type == "number":
                mistakes_found["inconsistent_numbering"]          = type_metrics.get("inconsistent_numbering", "no")
                mistakes_found["inconsistent_numbering_desc"]     = type_metrics.get("numbering_desc", "")
                mistakes_found["inconsistent_decimal_places"]     = type_metrics.get("inconsistent_decimal_places", "no")
                mistakes_found["inconsistent_decimal_places_desc"]= type_metrics.get("decimal_desc", "")
            elif detected_type == "date":
                mistakes_found["inconsistent_dates_formatting"]      = type_metrics.get("inconsistent_date_formatting", "no")
                mistakes_found["inconsistent_dates_formatting_desc"] = type_metrics.get("desc", "")
            elif detected_type == "phone":
                mistakes_found["missing_leading_zeros"]      = type_metrics.get("missing_leading_zeros", "no")
                mistakes_found["missing_leading_zeros_desc"] = type_metrics.get("desc", "")
            elif detected_type == "email":
                mistakes_found["invalid_emails"]       = type_metrics.get("invalid_emails", "no")
                mistakes_found["invalid_emails_desc"]  = type_metrics.get("invalid_emails_desc", "")
                mistakes_found["invalid_email_list"]   = type_metrics.get("invalid_email_list", [])
                mistakes_found["mixed_case_emails"]    = type_metrics.get("mixed_case_emails", "no")
                mistakes_found["mixed_case_emails_desc"] = type_metrics.get("mixed_case_emails_desc", "")
            elif detected_type == "text":
                mistakes_found["inconsistent_formatting"]      = type_metrics.get("inconsistent_formatting", "no")
                mistakes_found["inconsistent_formatting_desc"] = type_metrics.get("casing_desc", "")
                # NEW: full typo objects with suggestions
                mistakes_found["misspellings"] = typos

            column_diagnostics[col] = {
                "class": detected_type,
                "mistakes_found": mistakes_found
            }

        # ── Build suggested-clean rows ──────────────────────────────────────
        # Each cell gets the auto-suggested cleaned value.
        # The frontend will display these as the "preview" the user can override.
        suggested_rows = []
        for _, row in df.iterrows():
            clean_row = {}
            for col in headers:
                raw_val = row[col]
                cd      = column_diagnostics[col]
                _, tm   = analyze_exclusive_type(
                    df[col].dropna().astype(str).str.strip()
                )
                clean_row[col] = suggest_clean_value(
                    raw_val,
                    cd["class"],
                    tm,
                    spell_maps[col]
                )
            suggested_rows.append(clean_row)

        # Legacy pipe-joined rows (kept for backward compat)
        df_cleaned = df.fillna("")
        clean_rows = ["|".join([str(val) for val in row]) for _, row in df_cleaned.iterrows()]

        payload = {
            "headers":                headers,
            "rows_json":              clean_rows,           # raw original, pipe-joined
            "suggested_rows":         suggested_rows,       # NEW: [{col: cleaned_val, …}, …]
            "layout_alignment_errors": layout_shifts,
            "duplicate_row_indices":  duplicate_indices,
            "diagnostics":            column_diagnostics
        }

        return jsonify({**payload, "raw_json": json.dumps(payload)}), 200

    except Exception as e:
        return jsonify({"error": f"Internal MasterX workflow crash: {str(e)}"}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
