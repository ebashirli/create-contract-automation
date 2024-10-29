import os
import sys

import pandas as pd
from docx import Document
from docx2pdf import convert

# uv run --with pyinstaller -- pyinstaller --onefile main.py


def fill_conract(template_path, output_path, data):
    doc = Document(template_path)
    for paragraph in doc.paragraphs:
        for key, value in data.items():
            if key in paragraph.text:
                for run in paragraph.runs:
                    run.text = run.text.replace(f"[{key}]", str(value))
    doc.save(output_path)
    convert(output_path)


def generate_contract_from_log():
    arg1 = None if len(sys.argv) == 1 else sys.argv[1]
    BASE_LOC = arg1 if arg1 else os.path.dirname(os.getcwd())
    log_loc = os.path.join(BASE_LOC, "subcontractors log.xlsm")
    log = pd.ExcelFile(log_loc)

    sheet_names = log.sheet_names
    sheet_name_list = "\n".join([f"{i+1}. {x}" for i, x in enumerate(sheet_names)])

    sheet_no = int(input("Select sheet:\n" + sheet_name_list + "\n")) - 1
    df = pd.read_excel(log_loc, sheet_name=sheet_no)
    contract_no = int(input("Enter contract number: ")) - 1
    row = dict(df.iloc[contract_no])
    sheet_name = sheet_names[sheet_no]
    fill_conract(
        os.path.join(BASE_LOC, 'templates', f"{sheet_name}.docx"),
        os.path.join(BASE_LOC, str(contract_no + 1), "contract", "main.docx"),
        row,
    )


if __name__ == "__main__":
    generate_contract_from_log()

