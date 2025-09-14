from flask import Flask, render_template, request, redirect
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)
EXCEL_FILE = "contractor_data.xlsx"

# Define DIA columns
DIA_COLUMNS = [col.upper().strip() for col in [
    "63DIA", "75 DIA", "90 DIA", "110 DIA", "125 DIA", "140 DIA", "160 DIA", "180 DIA", "200 DIA"
]]

# Initialize Excel file with correct structure
if not os.path.exists(EXCEL_FILE):
    df = pd.DataFrame(columns=[
        "DATE", "RA BILL", "VENDOR CODE", "NAME OF THE CONTRACTOR", "SCHEME ID", "PANCHAYAT", "TYPE"
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
    contractor = request.form.get("contractorName", '').strip()
    vendor_code = request.form.get("vendorCode", '').strip()
    scheme_id = request.form.get("SchemeID", '').strip()
    panchayat = request.form.get("panchayat", '').strip()
    ra_bill = request.form.get("raBill", '').strip()
    date = request.form.get("workDate", '').strip()

    try:
        formatted_date = datetime.strptime(date, "%Y-%m-%d").strftime("%d-%m-%Y")
    except ValueError:
        formatted_date = ""

    # Load existing Excel data
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE)
        df.columns = df.columns.str.strip().str.upper()
    else:
        df = pd.DataFrame(columns=[
            "DATE", "RA BILL", "VENDOR CODE", "NAME OF THE CONTRACTOR", "SCHEME ID", "PANCHAYAT", "TYPE"
        ] + DIA_COLUMNS)

    # Prepare the new row
    bill_row = {
        "DATE": formatted_date,
        "RA BILL": ra_bill,
        "VENDOR CODE": vendor_code,
        "NAME OF THE CONTRACTOR": contractor,
        "SCHEME ID": scheme_id,
        "PANCHAYAT": panchayat,
        "TYPE": "this_bill"
    }

    # Fill DIA values
    for dia_label in DIA_COLUMNS:
        dia_value = dia_label.split()[0]
        bill_input = request.form.get(f"bill_{dia_value}", "0").strip()

        try:
            bill = int(bill_input) if bill_input else 0
        except ValueError:
            bill = 0

        bill_row[dia_label] = bill

    # Append and save
    df = pd.concat([df, pd.DataFrame([bill_row])], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)

    return redirect("/")

if __name__ == "__main__":
    app.run(debug=True)
