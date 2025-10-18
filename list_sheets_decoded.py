
import openpyxl
import sys

try:
    workbook = openpyxl.load_workbook("C:\InsuranceProject\main.xlsm")
    with open("C:\InsuranceProject\sheet_names_decoded.txt", "w", encoding="cp949") as f:
        for sheet_name in workbook.sheetnames:
            f.write(sheet_name + "\n")
except Exception as e:
    print(f"Error: {e}", file=sys.stderr)

