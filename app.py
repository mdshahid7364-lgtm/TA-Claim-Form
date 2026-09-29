from flask import Flask, render_template, request, send_from_directory
from openpyxl import load_workbook
import os
import shutil

app = Flask(__name__)


# Home page
@app.route("/")
def home():
    return render_template("form.html")


# TA Form Submit
@app.route("/submit", methods=["POST"])
def submit():

    # Form se employee data lena
    employee_code = request.form.get("employee_code")
    name = request.form.get("name")
    position = request.form.get("position")
    job = request.form.get("job")
    grade = request.form.get("grade")

    # Advance amount
    advance = request.form.get("advance")
    advance = float(advance) if advance else 0

    # Travel details
    travel_date = request.form.get("travel_date")
    location_from = request.form.get("location_from")
    location_to = request.form.get("location_to")
    from_time = request.form.get("from_time")
    to_time = request.form.get("to_time")

    # Expense details
    expense_dates = request.form.getlist("expense_date[]")
    particulars = request.form.getlist("particulars[]")
    modes = request.form.getlist("mode[]")
    supporting = request.form.getlist("supporting[]")
    actual_amounts = request.form.getlist("actual_amount[]")
    trs_policies = request.form.getlist("trs_policy[]")
    remarks = request.form.getlist("remarks[]")

    # Original Excel template
    template_file = "TA_Template.xlsx"

    # Generated folder banana
    os.makedirs("generated", exist_ok=True)

    # Employee ke naam se output file banana
    safe_name = name.replace(" ", "_")
    output_file = f"generated/TA_{safe_name}_{employee_code}.xlsx"

    # Original template ki copy banana
    shutil.copy(template_file, output_file)

    # Excel file open karna
    workbook = load_workbook(output_file)
    sheet = workbook.active

    # Employee details Excel me bharna
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

    # Expense table
    # Excel me first expense row 11 hai
    start_row = 11

    for i in range(len(expense_dates)):

        row = start_row + i

        sheet[f"A{row}"] = expense_dates[i]
        sheet[f"B{row}"] = particulars[i]
        sheet[f"C{row}"] = modes[i]
        sheet[f"D{row}"] = supporting[i]

        sheet[f"E{row}"] = (
            float(actual_amounts[i])
            if actual_amounts[i]
            else 0
        )

        sheet[f"F{row}"] = (
            float(trs_policies[i])
            if trs_policies[i]
            else 0
        )

        sheet[f"G{row}"] = remarks[i]

    # Total Actual Expense
    sheet["E23"] = "=SUM(E11:E22)"

    # Advance
    sheet["B25"] = advance

    # Balance
    sheet["B26"] = "=E23-B25"

    # Excel save karna
    workbook.save(output_file)

    # Sirf filename nikalna
    filename = os.path.basename(output_file)

    # Success page + Download button
    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>TA Submitted</title>
    </head>

    <body>

        <h1>TA Form Submitted Successfully!</h1>

        <p><b>Employee Code:</b> {employee_code}</p>
        <p><b>Name:</b> {name}</p>
        <p><b>Job:</b> {job}</p>
        <p><b>From:</b> {location_from}</p>
        <p><b>To:</b> {location_to}</p>

        <br>

        <p>
            <b>TA Excel file has been created successfully.</b>
        </p>

        <br>

        <a href="/download/{filename}">
            <button>
                Download TA Excel
            </button>
        </a>

    </body>
    </html>
    """


# Excel Download Route
@app.route("/download/<filename>")
def download_file(filename):

    return send_from_directory(
        "generated",
        filename,
        as_attachment=True
    )


# Flask server
if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=5000
    )
