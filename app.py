from flask import Flask, render_template, request
import pandas as pd
import os
import tempfile

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():
    output_lines = None
    error_msg = None
    cost_missing_rows = None

    if request.method == "POST":
        file = request.files.get("excel_file")
        columns = request.form.get("required_columns", "")
        # User does NOT enter COST; always add it for backend processing
        user_columns = [c.strip() for c in columns.split(",") if c.strip()]
        # Ensure COST is included
        required_columns = user_columns.copy()
        if "COST" not in [c.upper() for c in user_columns]:
            required_columns.append("COST")

        if not file or not user_columns:
            error_msg = "Please upload a file and specify parameter columns."
        else:
            temp_path = os.path.join(tempfile.gettempdir(), file.filename)
            file.save(temp_path)
            try:
                df = pd.read_excel(temp_path)
                missing_cols = [col for col in required_columns if col not in df.columns]
                if missing_cols:
                    error_msg = f"Missing columns in Excel file: {missing_cols}"
                else:
                    # Identify rows with missing COST values for user feedback
                    cost_missing_rows = [
                        idx + 2  # Excel-style (1-based, +header row)
                        for idx, val in df["COST"].items()
                        if pd.isna(val) or str(val).strip() == "" or val is None
                    ]
                    if not cost_missing_rows:
                        # Map column types: text or numeric
                        col_types = {}
                        for col in required_columns:
                            series = df[col]
                            if col.upper() == "COST":
                                col_types[col] = "numeric"
                            elif pd.api.types.is_numeric_dtype(series):
                                col_types[col] = "numeric"
                            else:
                                col_types[col] = "text"
                        lines = []
                        for _, row in df.iterrows():
                            clauses = []
                            # All columns except COST for 'WHEN ... AND ...'
                            for col in required_columns[:-1]:
                                col_name_quoted = f'"{col}"'
                                val = row[col]
                                if col_types[col] == "numeric":
                                    try:
                                        val_num = float(str(val).replace('$', '').replace(',', ''))
                                        val_fmt = f"{val_num}"
                                    except Exception:
                                        val_fmt = val
                                    formatted = f'{col_name_quoted} {val_fmt}'
                                else:
                                    formatted = f"{col_name_quoted} '{val}'"
                                clauses.append(formatted)
                            # COST column for THEN statement
                            cost_val = row["COST"]
                            try:
                                cost_val = float(str(cost_val).replace('$', '').replace(',', ''))
                                cost_fmt = f"{cost_val:.4f}"
                            except Exception:
                                cost_fmt = str(cost_val)
                            cost_clause = f'"COST" {cost_fmt}'
                            # Compose output line
                            when_and = " AND ".join(clauses)
                            output_line = f'WHEN {when_and} THEN {cost_clause}'
                            lines.append(output_line)
                        output_lines = lines
            except Exception as e:
                error_msg = str(e)
            finally:
                os.remove(temp_path)

    return render_template(
        "index.html",
        output_lines=output_lines,
        error_msg=error_msg,
        cost_missing_rows=cost_missing_rows
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
