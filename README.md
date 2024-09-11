# Excel To Outlook

This is a simple script that reads an Excel file that is structured like the provided `excel_to_outlook.xlsx` and adds the appointments to your Outlook calendar.
If the `Empfänger` column is filled, the appointment will be sent to the recipient.

Only the entries `Betreff` , `Start Datum`, `Start Uhrzeit` and `Dauer` are required. The other columns are optional.

## Installation
Create virtual environment and install the requirements:
```sh
python -m venv .venv
/.venv/Scripts/activate
pip install -r requirements.txt
```

## Usage
```sh
python excel_to_outlook.py <path_to_excel_file.xlsx>
```