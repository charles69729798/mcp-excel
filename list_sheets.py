
import openpyxl
import sys

try:
    workbook = openpyxl.load_workbook("C:\InsuranceProject\main.xlsm")
    for sheet_name in workbook.sheetnames:
        print(sheet_name)
except Exception as e:
    print(f"Error: {e}", file=sys.stderr)
