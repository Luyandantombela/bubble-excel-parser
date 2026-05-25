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

        df.columns = [str(col) for col in df.columns]
        headers = list(df.columns)
        df_cleaned = df.fillna("")

        # --- THE SIMPLE SEPARATOR ---
        # Joins every single cell value in the spreadsheet using a simple '|'
        all_cells = []
        for _, row in df_cleaned.iterrows():
            for val in row:
                all_cells.append(str(val))

        return (
            jsonify({"headers": headers, "rows_json": "|".join(all_cells)}),
            200,
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
