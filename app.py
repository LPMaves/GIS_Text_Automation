from flask import Flask, render_template, request
import pandas as pd
import os
import tempfile

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():
    output_lines = None
    error_msg = None

    if request.method == "POST":
        file = request.files.get("excel_file")
        columns = request.form.get("required_columns", "")
        required_columns = [c.strip() for c in columns.split(",") if c.strip()]

        if not file or not required_columns:
            error_msg = "Please upload a file and specify column names."
        else:
            temp_path = os.path.join(tempfile.gettempdir(), file.filename)
            file.save(temp_path)
            try:
                df = pd.read_excel(temp_path)
                missing_cols = [col for col in required_columns if col not in df.columns]
                if missing_cols:
                    error_msg = f"Missing columns in Excel file: {missing_cols}"
                else:
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
                        for idx, col in enumerate(required_columns):
                            col_name_quoted = f'"{col}"'
                            val = row[col]
                            if col_types[col] == "numeric":
                                if isinstance(val, str):
                                    val = val.replace('$', '').replace(',', '')
                                try:
                                    val_num = float(val)
                                    if col.upper() == "COST":
                                        val_fmt = f"{val_num:.4f}"
                                    else:
                                        val_fmt = f"{val_num}"
                                except Exception:
                                    val_fmt = val
                                formatted = f'{col_name_quoted} {val_fmt}'
                            else:
                                formatted = f"{col_name_quoted} '{val}'"
                            clauses.append(formatted)
                        if len(clauses) >= 2:
                            when_and = " AND ".join(clauses[:-1])
                            then_clause = clauses[-1]
                            output_line = f'WHEN {when_and} THEN {then_clause}'
                        else:
                            output_line = f'WHEN {clauses[0]}'
                        lines.append(output_line)
                    output_lines = lines
            except Exception as e:
                error_msg = str(e)
            finally:
                os.remove(temp_path)

    return render_template("index.html", output_lines=output_lines, error_msg=error_msg)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
