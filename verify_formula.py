
import openpyxl

FILE_PATH = "C:\\InsuranceProject\\신한주니어큐브종합건강상해보험(무배당, 해약환급금 미지급형)_주보험.xlsx"
SHEET_NAME = "기수표"

def verify_formula():
    """특정 셀의 수식을 확인합니다."""
    try:
        workbook = openpyxl.load_workbook(FILE_PATH, data_only=False)
        sheet = workbook[SHEET_NAME]
        cell_value = sheet['C2'].value
        with open("C:\\InsuranceProject\\formula_verify.txt", "w", encoding="utf-8") as f:
            f.write(str(cell_value))
    except Exception as e:
        with open("C:\\InsuranceProject\\formula_verify.txt", "w", encoding="utf-8") as f:
            f.write(f"오류: {e}")

if __name__ == "__main__":
    verify_formula()
