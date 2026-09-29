from flask import Flask, render_template, request
from openpyxl import load_workbook
import os
import shutil

app = Flask(__name__)


@app.route("/")
def home():
    return render_template("form.html")


@app.route("/submit", methods=["POST"])
def submit():

    # Form se data lena
    employee_code = request.form.get("employee_code")
    name = request.form.get("name")
    position = request.form.get("position")
    job = request.form.get("job")
    grade = request.form.get("grade")

    advance = request.form.get("advance")
    advance = float(advance) if advance else 0

    travel_date = request.form.get("travel_date")
    location_from = request.form.get("location_from")
    location_to = request.form.get("location_to")
    from_time = request.form.get("from_time")
    to_time = request.form.get("to_time")
        # Expense details receive karna
    expense_dates = request.form.getlist("expense_date[]")
    particulars = request.form.getlist("particulars[]")
    modes = request.form.getlist("mode[]")
    supporting = request.form.getlist("supporting[]")
    actual_amounts = request.form.getlist("actual_amount[]")
    trs_policies = request.form.getlist("trs_policy[]")
    remarks = request.form.getlist("remarks[]")

    # Template file
    template_file = "TA_Template.xlsx"

    # Generated folder banana
    os.makedirs("generated", exist_ok=True)

    # Employee ke naam se output file banana
    safe_name = name.replace(" ", "_")
    output_file = f"generated/TA_{safe_name}_{employee_code}.xlsx"

    # Original template ki copy banana
    shutil.copy(template_file, output_file)

    # Copy ko open karna
    workbook = load_workbook(output_file)
    sheet = workbook.active

    # TA format ke cells me data bharna
    sheet["B4"] = employee_code
    sheet["D4"] = travel_date

    sheet["G4"] = ""

    sheet["B5"] = name
    sheet["D5"] = location_from
    sheet["G5"] = from_time

    sheet["B6"] = position
    sheet["D6"] = location_to
    sheet["G6"] = to_time

    sheet["B7"] = job
    sheet["D7"] = grade

        # Expense table Excel me fill karna

    start_row = 9

    for i in range(len(expense_dates)):

        row = start_row + i

        sheet[f"A{row}"] = expense_dates[i]
        sheet[f"B{row}"] = particulars[i]
        sheet[f"C{row}"] = modes[i]
        sheet[f"D{row}"] = supporting[i]
        sheet[f"E{row}"] = float(actual_amounts[i]) if actual_amounts[i] else 0
        sheet[f"F{row}"] = float(trs_policies[i]) if trs_policies[i] else 0
        sheet[f"G{row}"] = remarks[i]

        sheet["E23"] = "=SUM(E9:E22)"
        sheet["B25"] = advance
        sheet["B26"] = "=E23-B25"
        # Excel save karna
    workbook.save(output_file)

    return f"""
    <h1>TA Form Submitted Successfully!</h1>

    <p>Employee Code: {employee_code}</p>
    <p>Name: {name}</p>
    <p>Job: {job}</p>
    <p>From: {location_from}</p>
    <p>To: {location_to}</p>

    <p><b>TA Excel file has been created successfully.</b></p>

    <p>File: {output_file}</p>
    """


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)