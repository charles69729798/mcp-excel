import time
import win32com.client
import os

FILE_PATH = os.path.abspath(r"C:\InsuranceProject\신한주니어큐브종합건강상해보험(무배당, 해약환급금 미지급형)_주보험_vba.xlsm")
MACRO_NAME = "Module2.RunAll"

def run_main_macro_debug():
    excel = None
    workbook = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True   # 잠시 True로 해보세요.
        
        print(f"Opening workbook: {FILE_PATH}")
        workbook = excel.Workbooks.Open(FILE_PATH)
        print("Workbook opened. Waiting 2 seconds...")
        time.sleep(2)          # 파일 완전 오픈 대기
        
        print(f"Attempting to run macro: {MACRO_NAME}")
        excel.Application.Run(MACRO_NAME)  # 또는 "파일명!모듈명.RunAll"
        print("Macro execution attempted.")

        workbook.Close(False)
        print("Workbook closed.")
        excel.Quit()
        print("Excel quit.")

    except Exception as e:
        print(f"오류 발생: {e}")
    finally:
        if workbook:
            workbook.Close(False)
        if excel:
            excel.Quit()

if __name__ == "__main__":
    run_main_macro_debug()
