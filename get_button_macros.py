
import win32com.client
import os

SOURCE_FILE_PATH = os.path.abspath("C:\\InsuranceProject\\034_프로그램_신한주니어큐브종합건강상해보험(무배당, 해약환급금 미지급형)_240401.xlsm")
OUTPUT_FILE = "C:\\InsuranceProject\\button_macros.txt"

def get_button_macros():
    """'S산출' 시트의 버튼에 연결된 매크로 이름을 가져옵니다."""
    excel = None
    workbook = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        workbook = excel.Workbooks.Open(SOURCE_FILE_PATH)

        with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
            f.write("--- Button Macro Analysis ---\\n")
            try:
                sheet = workbook.Sheets("S산출")
                for shape in sheet.Shapes:
                    if "Button" in shape.Name:
                        f.write(f"Button Name: {shape.Name}, Macro: {shape.OnAction}\\n")
            except Exception as e:
                f.write(f"Could not analyze buttons: {e}\\n")

        print("버튼 매크로 분석이 완료되었습니다.")

    except Exception as e:
        print(f"오류 발생: {e}")
    finally:
        if workbook:
            workbook.Close(False)
        if excel:
            excel.Quit()

if __name__ == "__main__":
    get_button_macros()
