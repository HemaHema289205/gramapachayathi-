from flask import Flask, render_template, request, redirect
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)
EXCEL_FILE = "contractor_data.xlsx"

DIA_COLUMNS = [col.upper().strip() for col in ["75 DIA", "90 DIA", "120 DIA", "140 DIA", "150 DIA", "180 DIA", "200 DIA", "220 DIA"]]

# Initialize Excel with correct structure
if not os.path.exists(EXCEL_FILE):
    df = pd.DataFrame(columns=[
        "DATE", "VENDOR CODE", "NAME OF THE CONTRACTOR", "SCHEME ID", "PANCHAYAT", "TYPE"
    ] + DIA_COLUMNS)
    df.to_excel(EXCEL_FILE, index=False)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/details")
def details():
    return render_template("details.html")

@app.route("/submit", methods=["POST"])
def submit():
    contractor = request.form.get("contractorName", '')
    vendor_code = request.form.get("vendorCode", '')
    scheme_id = request.form.get("SchemeID", '')
    panchayat = request.form.get("panchayat", '')
    date = request.form.get("workDate", '')
    formatted_date = datetime.strptime(date, "%Y-%m-%d").strftime("%d-%m-%Y")

    # Read the existing Excel file
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE)
        df.columns = df.columns.str.strip().str.upper()
    else:
        df = pd.DataFrame(columns=[
            "DATE", "VENDOR CODE", "NAME OF THE CONTRACTOR", "SCHEME ID", "PANCHAYAT", "TYPE"
        ] + DIA_COLUMNS)

    # Initialize a list to hold the new rows
    new_rows = []

    # Get data from the form for all DIA values
    form_data = {}
    for dia_label in DIA_COLUMNS:
        dia_value = dia_label.split()[0]
        bill_input = request.form.get(f"bill_{dia_value}", "0").strip()
        cum_input = request.form.get(f"cum_{dia_value}", "0").strip()

        try:
            bill = int(bill_input) if bill_input else 0
        except ValueError:
            bill = 0

        try:
            cum_value = int(cum_input) if cum_input else 0
        except ValueError:
            cum_value = 0

        form_data[dia_label] = {
            "bill": bill,
            "cum_value": cum_value
        }

    # Prepare the new rows
    cumulative_row = {
        "DATE": formatted_date,
        "VENDOR CODE": vendor_code,
        "NAME OF THE CONTRACTOR": contractor,
        "SCHEME ID": scheme_id,
        "PANCHAYAT": panchayat,
        "TYPE": "cumulative_sum"
    }

    previous_row = {
        "DATE": "",
        "VENDOR CODE": "",
        "NAME OF THE CONTRACTOR":"",
        "SCHEME ID": "",
        "PANCHAYAT": "",
        "TYPE": "previous_sum"
    }
    
    bill_row = {
        "DATE": "",
        "VENDOR CODE": "",
        "NAME OF THE CONTRACTOR": "",
        "SCHEME ID": "",
        "PANCHAYAT": "",
        "TYPE": "this_bill"
    }
    
    # Fill the rows with the correct DIA values
    for dia_label in DIA_COLUMNS:
        dia_data = form_data[dia_label]
        cumulative_row[dia_label] = dia_data["cum_value"]
        bill_row[dia_label] = dia_data["bill"]
        previous_row[dia_label] = dia_data["cum_value"] - dia_data["bill"]

    new_rows.append(cumulative_row)
    new_rows.append(previous_row)
    new_rows.append(bill_row)

    # Append the new rows to the DataFrame
    df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)

    # Save the updated DataFrame to the Excel file
    df.to_excel(EXCEL_FILE, index=False)
    
    return redirect("/")

if __name__ == "__main__":
    app.run(debug=True)
