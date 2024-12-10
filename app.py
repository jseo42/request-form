from flask import Flask, render_template, request, redirect
import pandas as pd
from pathlib import Path

app = Flask(__name__)

EXCEL_FILE = "requests.xlsx"

# Ensure the Excel file exists and is valid
def ensure_excel_file():
    if not Path(EXCEL_FILE).exists() or not Path(EXCEL_FILE).stat().st_size:
        df = pd.DataFrame(columns=["Name", "Request Type", "Description", "Priority", "Date"])
        df.to_excel(EXCEL_FILE, index=False, engine="openpyxl")

@app.route("/", methods=["GET", "POST"])
def index():
    ensure_excel_file()
    if request.method == "POST":
        name = request.form["name"]
        req_type = request.form["req_type"]
        description = request.form["description"]
        priority = request.form["priority"]

        # Read the existing Excel file
        df = pd.read_excel(EXCEL_FILE, engine="openpyxl")

        # Create a DataFrame for the new row
        new_row = pd.DataFrame([{
            "Name": name,
            "Request Type": req_type,
            "Description": description,
            "Priority": priority,
            "Date": pd.Timestamp.now()
        }])

        # Concatenate the new row with the existing DataFrame
        df = pd.concat([df, new_row], ignore_index=True)

        # Save back to the Excel file
        df.to_excel(EXCEL_FILE, index=False, engine="openpyxl")

        return redirect("/")

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)
