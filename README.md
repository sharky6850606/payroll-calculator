# Local Payroll Calculator

This Streamlit app calculates daily and total payroll hours from a biometric attendance Excel file.

The app processes only the worksheet named `Original record`. It scans repeated employee sections, reads clock-in and clock-out times, calculates worked hours in Python, writes the results into the green summary row for each employee, and creates a download file that contains only the updated `Original record` sheet.

Supported upload formats include modern Excel files such as `.xlsx` and `.xlsm`, plus legacy biometric exports such as Excel 97-2003 `.xls`. Legacy formats are converted to `.xlsx` before processing.

## Install

```bash
pip install -r requirements.txt
```

For legacy `.xls`, `.xlt`, `.xlsb`, and `.ods` uploads, install LibreOffice locally so the app can convert them:

```bash
winget install TheDocumentFoundation.LibreOffice
```

## Run Locally

```bash
streamlit run app.py
```

The app runs locally on your computer. It does not use a database, login, cloud storage, deployment, or API keys.

## Use The App

1. Open the Streamlit app in your browser.
2. Click **Upload your biometric Excel file**.
3. Select the workbook exported from the biometric attendance system.
4. Wait while the app processes `Original record`.
5. Click **Download updated Original record**.

The downloaded workbook is named `payroll_original_record_updated.xlsx`.

## Deploy On Streamlit Community Cloud

To share this app with the payroll team, put these files in a GitHub repository:

- `app.py`
- `requirements.txt`
- `packages.txt`
- `README.md`

Then go to Streamlit Community Cloud, create a new app, choose the GitHub repository, choose the branch, and set the app file to `app.py`.

`packages.txt` tells Streamlit Community Cloud to install LibreOffice so Excel 97-2003 `.xls` biometric exports can be converted automatically.

Important: once deployed to Streamlit Community Cloud, uploaded payroll workbooks are processed on Streamlit's hosted servers instead of only on your local computer. Do not upload sensitive payroll files there unless that is acceptable for your organization.

## Calculation Logic

Times are found with a regular expression, including values such as `08:45`, `8:45`, `13:16`, and `17:21`.

For each day, times are treated as pairs:

- 1st time: clock in
- 2nd time: clock out
- 3rd time: clock in
- 4th time: clock out

Incomplete pairs are ignored. If a clock-out time is earlier than its clock-in time, the app treats it as an overnight shift and adds 24 hours.
