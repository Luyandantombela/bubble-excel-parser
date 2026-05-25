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

        # Read file data directly into memory
        file_bytes = file.read()

        # Try reading as a modern Excel file (.xlsx) first
        try:
            df = pd.read_excel(file_bytes, engine="openpyxl")
        except Exception:
            # If that fails, fall back to the older Excel format engine (.xls)
            try:
                df = pd.read_excel(file_bytes, engine="xlrd")
            except Exception as engine_error:
                return (
                    jsonify(
                        {
                            "error": f"Could not determine Excel format. Details: {str(engine_error)}"
                        }
                    ),
                    400,
                )

        # Extract headers and clean rows
        headers = [str(col) for col in df.columns]
        df_cleaned = df.fillna("")
        rows_list = df_cleaned.to_dict(orient="records")

        return (
            jsonify({"headers": headers, "rows_json": json.dumps(rows_list)}),
            200,
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
