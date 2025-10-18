import win32com.client
import os

FILE_PATH = os.path.abspath("C:\\InsuranceProject\\신한주니어큐브종합건강상해보험(무배당, 해약환급금 미지급형)_주보험.xlsx")
SHEET_NAME = "기수표"
OUTPUT_FILE = "C:\\InsuranceProject\\formula_verify_win32.txt"

def verify_formula_win32():
    """특정 셀의 수식을 win32com을 사용하여 확인합니다."""
    excel = None
    workbook = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        workbook = excel.Workbooks.Open(FILE_PATH)
        sheet = workbook.Sheets(SHEET_NAME)
        
        cell_c2_formula = sheet.Range("C2").Formula
        cell_c2_formula_local = sheet.Range("C2").FormulaLocal
        
        cell_j1_formula = sheet.Range("J1").Formula
        cell_j1_formula_local = sheet.Range("J1").FormulaLocal
        
        with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
            f.write(f"C2 Formula: {cell_c2_formula}\n")
            f.write(f"C2 FormulaLocal: {cell_c2_formula_local}\n")
            f.write(f"J1 Formula: {cell_j1_formula}\n")
            f.write(f"J1 FormulaLocal: {cell_j1_formula_local}\n")

    except Exception as e:
        with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
            f.write(f"오류: {e}")
    finally:
        if workbook:
            workbook.Close(False)
        if excel:
            excel.Quit()

if __name__ == "__main__":
    verify_formula_win32()