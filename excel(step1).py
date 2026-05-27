import io, json, os, re, datetime
import pandas as pd
from flask import Flask, jsonify, request
from flask_cors import CORS

try:
    from spellchecker import SpellChecker
    # ── Instantiate ONCE at module load, not per column ──
    _SPELL = SpellChecker()
    SPELLCHECK_AVAILABLE = True
except ImportError:
    _SPELL = None
    SPELLCHECK_AVAILABLE = False

app = Flask(__name__)
CORS(app)

CUSTOM_WHITELIST = {
    'joburg', 'pretoria', 'durban', 'bongi', 'tina', 'chris', 'linda',
    'sarah', 'mike', 'amy', 'john', 'todo', 'asdf', 'idk'
}

# ── Pre-compile every regex used in hot loops ──
_RE_NON_ALPHA   = re.compile(r'[^a-zA-Z\s]')
_RE_NON_DIGIT_SEP = re.compile(r'[^\d\.\-]')
_RE_DATE1       = re.compile(r'^\d{4}[-/]\d{2}[-/]\d{2}$')
_RE_DATE2       = re.compile(r'^\d{2}[-/]\d{2}[-/]\d{2,4}$')
_RE_DATE_MASK   = re.compile(r'\d')
_RE_VALID_EMAIL = re.compile(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$')
_RE_NON_DIGIT   = re.compile(r'[\s\-\(\)\+]')
_RE_NUMBER_CLEAN= re.compile(r'[^\d\.,\-]')
_RE_CONTAMINATE = re.compile(r'[A-Za-z\$£€R]')
_RE_TITLE       = re.compile(r'^[A-Z][a-z]*(\s+[A-Z][a-z]*)*$')
_RE_NON_ALPHA_KEY = re.compile(r'[^a-zA-Z]')

_DATE_FORMATS = [
    '%Y/%m/%d', '%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y',
    '%m-%d-%y', '%m/%d/%Y', '%d %b %Y', '%b %d, %Y', '%m-%d-%Y'
]

_TEXT_NUMBERS = {
    'one','two','three','four','five','six','seven','eight','nine','ten',
    'twenty','thirty','forty','fifty'
}


# ─────────────────────────────────────────────
# SPELL CHECK
# ─────────────────────────────────────────────
def find_column_typos(series_str):
    if SPELLCHECK_AVAILABLE:
        found = {}
        for val in series_str:
            clean_text = _RE_NON_ALPHA.sub('', val)
            words = [w.lower() for w in clean_text.split() if len(w) > 2]
            unknown = _SPELL.unknown(words)
            for word in unknown:
                if word not in CUSTOM_WHITELIST and word not in found:
                    correction = _SPELL.correction(word)
                    found[word] = correction if correction and correction != word else None
        return [{"original": k, "suggestion": v} for k, v in found.items()]

    # Fallback: frequency heuristic — limit to top-200 words to avoid O(n²) explosion
    word_pool = []
    for val in series_str:
        clean_text = _RE_NON_ALPHA.sub('', val)
        for chunk in clean_text.split():
            if len(chunk) > 3:
                word_pool.append(chunk.lower())
    if not word_pool:
        return []

    freq_map = pd.Series(word_pool).value_counts()
    # Only compare top 200 words — beyond that the O(n²) loop is too expensive
    unique_words = list(freq_map.head(200).index)
    flagged = set()
    for i, word in enumerate(unique_words):
        for candidate in unique_words[i + 1:]:
            if abs(len(word) - len(candidate)) <= 2:
                if (word[:-1] == candidate or candidate[:-1] == word or
                        (word[:-2] == candidate[:-2] and word[:3] == candidate[:3])):
                    if not (word.startswith("unite") or candidate.startswith("unite")):
                        rare = word if freq_map[word] < freq_map[candidate] else candidate
                        flagged.add(rare)
    return [{"original": w, "suggestion": None} for w in flagged]


# ─────────────────────────────────────────────
# TYPE DETECTION  — operates on a plain Python list for speed
# ─────────────────────────────────────────────
def analyze_exclusive_type(values: list):
    """
    values: list of non-blank stripped strings (already filtered upstream).
    Returns (type_str, metrics_dict).
    """
    filled_count = len(values)
    if filled_count == 0:
        return "text", {}

    # DATE
    date_matches = sum(
        1 for x in values
        if _RE_DATE1.match(x) or _RE_DATE2.match(x)
    )
    if date_matches / filled_count > 0.4:
        unique_masks = len({_RE_DATE_MASK.sub('X', x) for x in values})
        has_mixed = "yes" if unique_masks > 1 else "no"
        return "date", {
            "inconsistent_date_formatting": has_mixed,
            "desc": "Mixed date formats found." if has_mixed == "yes" else "Dates are uniform."
        }

    # EMAIL
    at_count = sum(1 for x in values if '@' in x)
    if at_count / filled_count > 0.4:
        has_invalid, has_mixed_case = "no", "no"
        invalid_list = []
        for val in values:
            if not _RE_VALID_EMAIL.match(val):
                has_invalid = "yes"
                if val:
                    invalid_list.append(val)
            if any(c.isupper() for c in val):
                has_mixed_case = "yes"
        return "email", {
            "invalid_emails": has_invalid,
            "invalid_emails_desc": "Column contains broken email formats.",
            "invalid_email_list": list(set(invalid_list)),
            "mixed_case_emails": has_mixed_case,
            "mixed_case_emails_desc": (
                "Emails contain mixed uppercase letters."
                if has_mixed_case == "yes" else "Email casing is uniform."
            )
        }

    # PHONE
    phone_matches = sum(
        1 for x in values
        if (d := _RE_NON_DIGIT.sub('', x)).isdigit() and 7 <= len(d) <= 15
    )
    if phone_matches / filled_count > 0.4:
        has_issue, issue_desc = "no", "Phone numbers are uniform."
        for val in values:
            digits_only = re.sub(r'\D', '', val)
            if '?' in val or val.isalpha():
                has_issue = "yes"
                issue_desc = "Phone numbers contain invalid placeholder text symbols (like '??')."
                break
            elif val.startswith(('1','2','3','4','5','6','7','8','9')) and not val.startswith('+'):
                has_issue = "yes"
                issue_desc = "Phone numbers are truncated, missing leading zeros."
                break
            elif digits_only and len(digits_only) < 9:
                has_issue = "yes"
                issue_desc = "Phone numbers contain broken, short sequences missing digits."
                break
        return "phone", {"missing_leading_zeros": has_issue, "desc": issue_desc}

    # NUMBER
    number_score, has_contamination, decimal_lengths = 0, "no", []
    for val in values:
        cleaned = _RE_NUMBER_CLEAN.sub('', val.replace(',', '.'))
        if re.match(r'^-?\d+(\.\d+)?$', cleaned):
            number_score += 1
            if _RE_CONTAMINATE.search(val) or ('-' in val and not val.strip().startswith('-')):
                has_contamination = "yes"
            decimal_lengths.append(len(cleaned.split('.')[-1]) if '.' in cleaned else 0)
        elif val.lower() in _TEXT_NUMBERS:
            number_score += 1
            has_contamination = "yes"

    if number_score / filled_count > 0.4:
        inconsistent_decimals = "yes" if len(set(decimal_lengths)) > 1 else "no"
        return "number", {
            "inconsistent_numbering": has_contamination,
            "numbering_desc": (
                "Math cells contain currency symbols or text characters."
                if has_contamination == "yes" else "Numbers are cleanly formatted."
            ),
            "inconsistent_decimal_places": inconsistent_decimals,
            "decimal_desc": (
                "Uneven decimal lengths found across numbers."
                if inconsistent_decimals == "yes" else "Decimal lengths are uniform."
            )
        }

    # TEXT
    lower_count = upper_count = title_count = 0
    has_newlines = "no"
    for val in values:
        if "\n" in val or "\r" in val:
            has_newlines = "yes"
        if val.islower():
            lower_count += 1
        elif val.isupper():
            upper_count += 1
        elif _RE_TITLE.match(val):
            title_count += 1

    inconsistent_casing = (
        "no" if (lower_count == filled_count or
                 upper_count == filled_count or
                 title_count == filled_count)
        else "yes"
    )
    return "text", {
        "inconsistent_formatting": inconsistent_casing,
        "casing_desc": (
            "Hidden newline breaks found breaking cell format limits."
            if has_newlines == "yes" else "Inconsistent text casing layouts found."
        )
    }


# ─────────────────────────────────────────────
# SUGGEST CLEAN VALUE  (vectorised per column, not per cell)
# ─────────────────────────────────────────────
def suggest_clean_column(raw_series: pd.Series, detected_type: str,
                         type_metrics: dict, spell_map: dict) -> pd.Series:
    """
    Returns a cleaned Series for the whole column at once.
    Much faster than calling a per-cell function inside iterrows().
    """
    s = raw_series.fillna("").astype(str).str.strip()

    if detected_type == "text":
        def clean_text(val):
            if not val:
                return ""
            words = val.split()
            corrected = []
            for w in words:
                key = _RE_NON_ALPHA_KEY.sub('', w).lower()
                if key in spell_map and spell_map[key]:
                    sg = spell_map[key]
                    if w.isupper():       sg = sg.upper()
                    elif w.istitle():     sg = sg.title()
                    corrected.append(sg)
                else:
                    corrected.append(w)
            result = " ".join(corrected)
            if type_metrics.get("inconsistent_formatting") == "yes":
                result = result.title()
            return result
        return s.apply(clean_text)

    if detected_type == "number":
        def clean_number(val):
            if not val:
                return ""
            cleaned = _RE_NON_DIGIT_SEP.sub('', val.replace(',', '.'))
            try:
                return f"{float(cleaned):.2f}"
            except ValueError:
                return val
        return s.apply(clean_number)

    if detected_type == "date":
        def clean_date(val):
            if not val:
                return ""
            for fmt in _DATE_FORMATS:
                try:
                    return datetime.datetime.strptime(val, fmt).strftime('%Y/%m/%d')
                except ValueError:
                    continue
            return val
        return s.apply(clean_date)

    if detected_type == "email":
        return s.str.lower()

    if detected_type == "phone":
        def clean_phone(val):
            if not val:
                return ""
            digits = re.sub(r'\D', '', val)
            if digits and not val.startswith('+') and len(digits) >= 9:
                return '0' + digits if not digits.startswith('0') else digits
            return val
        return s.apply(clean_phone)

    return s  # passthrough for unknown types


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

        headers    = list(df.columns)
        total_rows = len(df)

        duplicate_mask    = df.duplicated(keep='first')
        duplicate_indices = [int(i + 1) for i, v in enumerate(duplicate_mask) if v]

        column_diagnostics = {}
        layout_shifts      = []
        spell_maps         = {}
        col_types          = {}   # col → detected_type
        col_metrics        = {}   # col → type_metrics

        # ── Single pass: diagnostics per column ──────────────────────────
        for i, col in enumerate(headers):
            series   = df[col]
            # Build plain list once — reused for type detection & spellcheck
            col_vals = (
                series.dropna()
                      .astype(str)
                      .str.strip()
                      .loc[lambda s: s != ""]
                      .tolist()
            )
            filled_count = len(col_vals)

            if total_rows > 5 and filled_count > 0 and (filled_count / total_rows) < 0.15:
                layout_shifts.append({
                    "column": col,
                    "error_msg": f"Stray text boundary shift in {col}.",
                    "sample_value": col_vals[0] if col_vals else ""
                })

            blank_count = int(
                series.isna().sum() +
                (series.astype(str).str.strip() == "").sum()
            )

            detected_type, type_metrics = analyze_exclusive_type(col_vals)
            col_types[col]   = detected_type
            col_metrics[col] = type_metrics

            typos      = []
            spell_map  = {}
            if detected_type == "text" and col_vals:
                typos      = find_column_typos(pd.Series(col_vals))
                spell_map  = {t["original"]: t["suggestion"] for t in typos}
            spell_maps[col] = spell_map

            mistakes_found = {
                "blank_cells":      blank_count,
                "blank_cells_desc": (
                    f"Found {blank_count} empty rows missing values."
                    if blank_count > 0 else "No missing values."
                )
            }

            if detected_type == "number":
                mistakes_found["inconsistent_numbering"]           = type_metrics.get("inconsistent_numbering", "no")
                mistakes_found["inconsistent_numbering_desc"]      = type_metrics.get("numbering_desc", "")
                mistakes_found["inconsistent_decimal_places"]      = type_metrics.get("inconsistent_decimal_places", "no")
                mistakes_found["inconsistent_decimal_places_desc"] = type_metrics.get("decimal_desc", "")
            elif detected_type == "date":
                mistakes_found["inconsistent_dates_formatting"]      = type_metrics.get("inconsistent_date_formatting", "no")
                mistakes_found["inconsistent_dates_formatting_desc"] = type_metrics.get("desc", "")
            elif detected_type == "phone":
                mistakes_found["missing_leading_zeros"]      = type_metrics.get("missing_leading_zeros", "no")
                mistakes_found["missing_leading_zeros_desc"] = type_metrics.get("desc", "")
            elif detected_type == "email":
                mistakes_found["invalid_emails"]        = type_metrics.get("invalid_emails", "no")
                mistakes_found["invalid_emails_desc"]   = type_metrics.get("invalid_emails_desc", "")
                mistakes_found["invalid_email_list"]    = type_metrics.get("invalid_email_list", [])
                mistakes_found["mixed_case_emails"]     = type_metrics.get("mixed_case_emails", "no")
                mistakes_found["mixed_case_emails_desc"]= type_metrics.get("mixed_case_emails_desc", "")
            elif detected_type == "text":
                mistakes_found["inconsistent_formatting"]      = type_metrics.get("inconsistent_formatting", "no")
                mistakes_found["inconsistent_formatting_desc"] = type_metrics.get("casing_desc", "")
                mistakes_found["misspellings"]                 = typos

            column_diagnostics[col] = {
                "class":          detected_type,
                "mistakes_found": mistakes_found
            }

        # ── Build suggested-clean rows (vectorised, one pass per column) ──
        cleaned_cols = {}
        for col in headers:
            cleaned_cols[col] = suggest_clean_column(
                df[col],
                col_types[col],
                col_metrics[col],
                spell_maps[col]
            ).tolist()

        # Transpose: list-of-dicts from dict-of-lists
        suggested_rows = [
            {col: cleaned_cols[col][ri] for col in headers}
            for ri in range(total_rows)
        ]

        # Legacy pipe-joined rows
        df_cleaned = df.fillna("")
        clean_rows = [
            "|".join(str(v) for v in row)
            for row in df_cleaned.itertuples(index=False, name=None)
        ]

        payload = {
            "headers":                 headers,
            "rows_json":               clean_rows,
            "suggested_rows":          suggested_rows,
            "layout_alignment_errors": layout_shifts,
            "duplicate_row_indices":   duplicate_indices,
            "diagnostics":             column_diagnostics
        }

        return jsonify({**payload, "raw_json": json.dumps(payload)}), 200

    except Exception as e:
        return jsonify({"error": f"Internal MasterX workflow crash: {str(e)}"}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
