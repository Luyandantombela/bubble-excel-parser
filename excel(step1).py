import json
import os
import pandas as pd
from flask import Flask, jsonify, request
from flask_cors import CORS

app = Flask(__name__)
CORS(app)  # Allows your Bubble app to fetch data safely


@app.route("/parse-excel", methods=["POST"])
def parse_excel():
    try:
        # Check if an Excel file was sent
        if "file" not in request.files:
            return jsonify({"error": "No file part in the request"}), 400

        file = request.files["file"]
        if file.filename == "":
            return jsonify({"error": "No selected file"}), 400

        # Read the Excel file directly from the upload stream
        df = pd.read_excel(file)

        # Extract headers and clean row data
        headers = [str(col) for col in df.columns]
        df_cleaned = df.fillna("")
        rows_list = df_cleaned.to_dict(orient="records")

        # Return structured data back to Bubble
        return (
            jsonify({"headers": headers, "rows_json": json.dumps(rows_list)}),
            200,
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
