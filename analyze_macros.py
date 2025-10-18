
import win32com.client
import os

SOURCE_FILE_PATH = os.path.abspath("C:\\InsuranceProject\\main.xlsm")
OUTPUT_FILE = "C:\\InsuranceProject\\macro_analysis.txt"

def analyze_macros():
    """'S산출' 시트의 매크로와 VBA 코드를 분석합니다."""
    excel = None
    workbook = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        workbook = excel.Workbooks.Open(SOURCE_FILE_PATH)

        with open(OUTPUT_FILE, "w", encoding="utf-8") as f_output:
            # 1. 버튼 매크로 분석
            f_output.write("--- Button Macro Analysis ---\n")
            try:
                sheet = workbook.Sheets("S산출")
                for shape in sheet.Shapes:
                    if "Button" in shape.Name:
                        f_output.write(f"Button Name: {shape.Name}, Macro: {shape.OnAction}\n")
            except Exception as e:
                f_output.write(f"Could not analyze buttons: {e}\n")

            # 2. VBA 코드 추출
            f_output.write("\n--- VBA Code Extraction ---\n")
            for component in workbook.VBProject.VBComponents:
                f_output.write(f"\n--- Module: {component.Name}, Type: {component.Type} ---\n")
                try:
                    if component.CodeModule.CountOfLines > 0:
                        f_output.write(component.CodeModule.Lines(1, component.CodeModule.CountOfLines))
                    else:
                        f_output.write("(No code in this module)\n")
                except Exception as e:
                    f_output.write(f"(Error reading code from module: {e})\n")

        print("매크로 분석이 완료되었습니다. 상세 내용은 analyze_macros_debug.txt 파일을 확인하세요.")

    except Exception as e:
        print(f"오류 발생: {e}")
    finally:
        if workbook:
            workbook.Close(False)
        # if excel:
        #     excel.Quit()

if __name__ == "__main__":
    analyze_macros()
