import io
import json
import os
import numpy as np
import pandas as pd
from flask import Flask, jsonify, request
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

def safe_str(val):
    if pd.isna(val):
        return ""
    if isinstance(val, (pd.Timestamp, np.datetime64)):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val).replace("|||", " ") # Prevent separator collisions

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
            df = pd.read_excel(file_stream, engine="openpyxl")
        except Exception:
            file_stream.seek(0)
            try:
                df = pd.read_excel(file_stream, engine="xlrd")
            except Exception as engine_error:
                return jsonify({"error": f"Format error: {str(engine_error)}"}), 400

        # Clean out header names
        df.columns = [str(col) if not str(col).startswith("Unnamed:") else f"Column {i+1}" for i, col in enumerate(df.columns)]
        headers = list(df.columns)

        # --- NEW BULLETPROOF FORMAT FOR BUBBLE ---
        # Converts each row into a simple string like: "Value1|||Value2|||Value3"
        formatted_rows = []
        for _, row in df.iterrows():
            row_string = "|||".join([safe_str(val) for val in row])
            formatted_rows.append(row_string)

        return jsonify({
            "headers": headers,
            "rows_json": "||||||".join(formatted_rows) # Glues rows together
        }), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
