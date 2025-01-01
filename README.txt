This is the README file that delineates everything that is happening in this application.

Directories and Files:
    excel_sheets:
        - 36Months.xlsx: Excel sheet displaying Unichem products that have a shelf life of 36 Months
        - Bayshore_Template.xlsm: Excel Template for Bayshore
        - Configurations.xlsx: Excel sheet displaying product configurations for all Unichem products
        - Pricing.xlsx: Excel sheet displaying pricing for all Unichem products
        - Unichem_Template.xlsm: Excel Template for Unichem

    json_files:
        - 36Months.json: json file for 36Months data created from create_json_files.py
        - configurations.json: json file for configurations data created from create_json_files.py
        - pricing.json: json file for pricing data created from create_json_files.py
        - json_data.json: json file written by read_bayshore_sheet.py or read_unichem.py and read by create_new_hda_file.py

    python_files:
        - constants.py: Python file containing a dictionary of key-value pairs mapping product_fields to Excel coordinates
        - create_json_files.py: Python file that reads in 36Months.xlsx, Configurations.xlsx, Pricing.xlsx and converts
            the date in these files to json files
        - create_new_hda_file: Python file that reads from json_data.json and populates a new 2024 HDA form with this data
        - read_bayshore_sheet: Python file that reads from an old HDA form and writes to json_data.json for Bayshore products
        - read_unichem_sheet: Python file that reads from an old HDA form and writes to json_data.json for Unichem products

    templates: (folder for creating HTMl files and should be left untouched)

    app.py: (intersection of frontend and backend code that allows user to make API calls to execute Python files)


Backend Workflow for HDA Form Conversions:
    1) Run create_json_files.py to create 36Months.json, Configurations.json, and Pricing.json. This step does NOT need
        to be run again, UNLESS there are new Excel sheets for 36Months.xlsx, Configurations.xlsx, and Pricing.xlsx.
    2) Run app.py to create web app.
        What is happening inside app.py?
        a) First, runs read_bayshore_sheet.py or read_unichem_sheet.py to read from uploaded Excel file and write to json_data.json
        b) Then, runs create_new_hda_file.py to write to a new Excel file.


Backend Edits for New HDA Form Conversions:



