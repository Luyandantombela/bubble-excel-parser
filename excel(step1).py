import io
import json
import os
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

        # 1. Clean and track the headers
        df.columns = [str(col) for col in df.columns]
        headers = list(df.columns)
        
        # Explicitly include the index column '0' as the very first header
        if "0" not in headers:
            headers.insert(0, "0")
            
        df_cleaned = df.fillna("")

        # --- THE ROW SEPARATOR FIX ---
        # Instead of mashing cells flat, we preserve the row lines
        clean_rows = []
        for index, row in df_cleaned.iterrows():
            # Add the row index number (0, 1, 2, 3...) to the front of each line
            row_values = [str(index + 1)] + [str(val) for val in row]
            row_string = "|".join(row_values)
            clean_rows.append(row_string)

        # We return rows_json as a list of full sentences instead of a flat string
        return jsonify({"headers": headers, "rows_json": clean_rows}), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

