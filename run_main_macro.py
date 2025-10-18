
import win32com.client
import os

FILE_PATH = os.path.abspath("C:\\InsuranceProject\\신한주니어큐브종합건강상해보험(무배당, 해약환급금 미지급형)_주보험_vba.xlsm")
MACRO_NAME = "RunAll"

def run_main_macro():
    """메인 매크로를 실행합니다."""
    excel = None
    workbook = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        workbook = excel.Workbooks.Open(FILE_PATH)

        # Run the macro
        excel.Application.Run(MACRO_NAME)

        workbook.Save()
        print(f"'{os.path.basename(FILE_PATH)}' 파일에서 매크로를 실행했습니다.")

    except Exception as e:
        print(f"오류 발생: {e}")
    finally:
        if workbook:
            workbook.Close(False)
        if excel:
            excel.Quit()

if __name__ == "__main__":
    run_main_macro()
