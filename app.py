from flask import Flask, render_template, send_file, jsonify, request
import time
from python_files.read_unichem_sheet import read_hda_sheet
from python_files.read_bayshore_sheet import read_bayshore_sheet
from python_files.create_new_hda_file import create_new_hda_file
import os

app = Flask(__name__)

most_recently_uploaded_file = [""]
most_recently_created_file = [""]


@app.route('/', methods=['GET', 'POST'])
def index():
    return render_template('index.html')


FOLDER = "excel_sheets"
if not os.path.exists(FOLDER):
    os.makedirs(FOLDER)

app.config["UPLOAD_FOLDER"] = FOLDER


@app.route("/upload", methods=["POST"])
def upload_file():
    directory_clean_up()
    # # Deletes previously uploaded Excel sheets.
    if most_recently_uploaded_file[0] != "":
        try:
            if os.path.exists(most_recently_uploaded_file[0]):
                os.remove(most_recently_uploaded_file[0])
        except Exception as e:
            print(f"An error occurred: {e}")
    if most_recently_created_file[0] != "":
        try:
            if os.path.exists(most_recently_created_file[0]):
                os.remove(most_recently_created_file[0])
        except Exception as e:
            print(f"An error occurred: {e}")
    if "excelFile" not in request.files:
        return jsonify({"error": "No file part"}), 400

    file = request.files["excelFile"]

    if file.filename == "":
        return jsonify({"error": "No selected file"}), 400

    filepath = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
    most_recently_uploaded_file[0] = filepath
    file.save(filepath)

    ################################################################## SCROLL TO HERE
    company = request.form.get("company")
    if not company:
        return "No company option selected", 400
    if company == 'UNICHEM':
        # writes to json_data.json to populate with Bayshore data
        read_hda_sheet(filepath[filepath.index('\\') + 1:])
    else:
        # writes to json_data.json to populate with Unichem data
        read_bayshore_sheet(filepath[filepath.index('\\') + 1:])
    time.sleep(.1)

    # reads from json_data.json, creates a new Unichem / Bayshore template and writes to it
    most_recently_created_file[0] = create_new_hda_file(company)
    ##################################################################
    return jsonify({"message": f"File successfully saved"}), 200


@app.route('/download')
def download_file():
    return send_file(most_recently_created_file[0], as_attachment=True)


STATIC_EXCEL_SHEETS = ["36Months.xlsx", "Bayshore_Template.xlsm", "Configurations.xlsx", "Pricing.xlsx", "Unichem_Template.xlsm"]


# deletes any old uploaded HDA if there are > 10 forms in the excel_sheets directory
def directory_clean_up():
    count = 0
    for _, _, files in os.walk(FOLDER):
        count += len(files)
    if count >= 15:
        directory = os.fsencode(FOLDER)

        copy = os.listdir(directory).copy()
        for file in copy:
            filename = os.fsdecode(file)
            if filename in STATIC_EXCEL_SHEETS:
                continue
            else:
                try:
                    if os.path.exists('excel_sheets\\' + filename):
                        os.remove('excel_sheets\\' + filename)
                except Exception as e:
                    print(f"An error occurred: {e}")


if __name__ == "__main__":
    app.run(debug=True)
