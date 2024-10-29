import os

import pandas as pd
from docx import Document
from docx2pdf import convert

# uv run --with pyinstaller -- pyinstaller main.py


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
    BASE_LOC = os.path.dirname(os.path.dirname(__file__))
    log_loc = os.path.join(BASE_LOC, "subcontractors log.xlsm")
    log = pd.ExcelFile(log_loc)

    sheet_names = "\n".join([f"{i+1}. {x}" for i, x in enumerate(log.sheet_names)])
    sheet_no = int(input("Select sheet:\n" + sheet_names + "\n")) - 1
    df = pd.read_excel(log_loc, sheet_name=sheet_no)
    contract_no = int(input("Enter contract number: "))
    row = dict(df.iloc[contract_no - 1])
    
    fill_conract(
        "temp.docx",
        os.path.join(BASE_LOC, str(contract_no), "contract", "main.docx"),
        row,
    )


if __name__ == "__main__":
    generate_contract_from_log()
