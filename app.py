import os
import tempfile
from flask import Flask, render_template, request
import pandas as pd

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():
    error_msg = None
    case_code = None
    cost_missing_rows = None

    if request.method == "POST":
        file = request.files.get("excel_file")
        columns_from_form = [
            c.strip() for c in request.form.get("required_columns", "") .split(",") if c.strip()
        ]
        layer_name = request.form.get("layer_name", "").strip()
        if not layer_name:
            error_msg = "Layer name is required."
        # Always append COST if not included (ignore case)
        upper_cols = [c.upper() for c in columns_from_form]
        if "COST" not in upper_cols:
            required_columns = columns_from_form + ["COST"]
        else:
            required_columns = columns_from_form

        # Get the subtraction value; default to 0
        subtract_raw = request.form.get("subtract_value", "")
        try:
            subtract_value = float(subtract_raw) if subtract_raw.strip() else 0.0
        except Exception:
            error_msg = f"Subtract value must be a number (got '{subtract_raw}')"
            subtract_value = 0.0

        if not file or not columns_from_form or not layer_name:
            if not error_msg:
                error_msg = "Please upload a file, specify parameter columns, and enter a layer name."
        else:
            temp_path = os.path.join(tempfile.gettempdir(), file.filename)
            file.save(temp_path)
            try:
                df = pd.read_excel(temp_path)
                # Validate column presence
                missing_cols = [col for col in required_columns if col not in df.columns]
                if missing_cols:
                    error_msg = f"Missing columns in Excel file: {missing_cols}"
                else:
                    # COST row-wise missing checks (using pandas NA/nan logic)
                    cost_missing_rows = [
                        idx + 2  # +2 for Excel-like row number (header is row 1)
                        for idx, val in df["COST"].items()
                        if pd.isna(val) or str(val).strip() == "" or val is None
                    ]

                    if not cost_missing_rows:
                        # Build CASE lines
                        case_lines = []
                        param_columns = [col for col in required_columns if col.upper() != "COST"]

                        for _, row in df.iterrows():
                            conditions = []
                            for col in param_columns:
                                col_fmt = f'"{col}"'
                                val = row[col]
                                if pd.api.types.is_numeric_dtype(df[col]):
                                    try:
                                        if pd.isna(val) or val is None or str(val).strip() == '':
                                            val_fmt = 'NULL'
                                        else:
                                            val_fmt = float(val)
                                    except Exception:
                                        val_fmt = val
                                    cond = f'{col_fmt} = {val_fmt}'
                                else:
                                    cond = f"{col_fmt} = '{val}'"
                                conditions.append(cond)
                            cost_val = row["COST"]
                            if isinstance(cost_val, str):
                                cost_val = cost_val.replace('$', '').replace(',', '')
                            try:
                                cost_float = float(cost_val)
                                cost_str = f"{cost_float:.2f}"
                            except Exception:
                                cost_str = str(cost_val)
                            line = f"WHEN " + " AND ".join(conditions) + f" THEN {cost_str}"
                            case_lines.append(line)

                        # Compose the full CASE block including the provided layer name
                        case_block = (
                            "'$' || format_number(\n"
                            f"  aggregate(\n"
                            f"      layer:='{layer_name}',\n"
                            "      aggregate:='sum',\n"
                            "      expression:=\n"
                            "          CASE\n"
                        )
                        for cl in case_lines:
                            case_block += f"              {cl}\n"
                        case_block += (
                            "              ELSE 0\n"
                            "          END\n"
                            "      ,2)\n"
                            f"  - {subtract_value:.4f}                        -- we are now subtracting the summation of the cost of the design\n"
                            ")\n)\n"
                        )

                        case_code = case_block

            except Exception as e:
                error_msg = f"Processing error: {e}"
            finally:
                if os.path.exists(temp_path):
                    os.remove(temp_path)

    return render_template(
        "index.html",
        error_msg=error_msg,
        case_code=case_code,
        cost_missing_rows=cost_missing_rows
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
