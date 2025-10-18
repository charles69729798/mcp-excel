
import win32com.client
import os

SOURCE_FILE_PATH = os.path.abspath("C:\\InsuranceProject\\main.xlsm")
TARGET_FILE_PATH = os.path.abspath("C:\\InsuranceProject\\신한주니어큐브종합건강상해보험(무배당, 해약환급금 미지급형)_주보험.xlsx")
NEW_TARGET_FILE_PATH = os.path.abspath("C:\\InsuranceProject\\신한주니어큐브종합건강상해보험(무배당, 해약환급금 미지급형)_주보험_vba.xlsm")

def replicate_s_sheet():
    """'S산출' 시트를 복사하고 VBA 코드를 주입합니다."""
    excel = None
    source_wb = None
    target_wb = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False

        # 1. 소스 및 대상 파일 열기
        source_wb = excel.Workbooks.Open(SOURCE_FILE_PATH)
        target_wb = excel.Workbooks.Open(TARGET_FILE_PATH)

        # 2. 'S산출' 시트 복사
        source_sheet = source_wb.Sheets("S산출")
        source_sheet.Copy(Before=target_wb.Sheets(1))
        # The copied sheet becomes the active sheet in the target workbook
        new_sheet = target_wb.ActiveSheet
        new_sheet.Name = "s산출"

        # Populate input cells
        for i in range(2, 14):
            new_sheet.Cells(3, i).Value = 1

        # 3. VBA 코드 주입
        with open("C:\\InsuranceProject\\vba_code_manual.txt", "r", encoding="utf-8") as f:
            full_vba_code = f.read()

        # Split the code into Sheet18 and Module2 parts
        sheet18_code_start = full_vba_code.find("--- Module: Sheet18, Type: 100 ---")
        module2_code_start = full_vba_code.find("--- Module: Module2, Type: 1 ---")

        sheet18_code = ""
        module2_code = ""

        if sheet18_code_start != -1:
            sheet18_code_end = full_vba_code.find("--- Module:", sheet18_code_start + 1)
            if module2_code_start != -1 and module2_code_start < sheet18_code_end:
                sheet18_code_end = module2_code_start
            if sheet18_code_end == -1:
                sheet18_code = full_vba_code[sheet18_code_start + len("--- Module: Sheet18, Type: 100 ---\n"):].strip()
            else:
                sheet18_code = full_vba_code[sheet18_code_start + len("--- Module: Sheet18, Type: 100 ---\n"):sheet18_code_end].strip()

        if module2_code_start != -1:
            module2_code_end = full_vba_code.find("--- Module:", module2_code_start + 1)
            if module2_code_end == -1:
                module2_code = full_vba_code[module2_code_start + len("--- Module: Module2, Type: 1 ---\n"):].strip()
            else:
                module2_code = full_vba_code[module2_code_start + len("--- Module: Module2, Type: 1 ---\n"):module2_code_end].strip()

        # 4. xlsm으로 저장
        target_wb.SaveAs(NEW_TARGET_FILE_PATH, FileFormat=52)
        print(f"파일이 '{os.path.basename(NEW_TARGET_FILE_PATH)}'으로 저장되었습니다.")

        # Re-open the new xlsm file to inject VBA
        target_wb.Close(False)
        target_wb = excel.Workbooks.Open(NEW_TARGET_FILE_PATH)
        new_sheet = target_wb.Sheets("s산출")

        # Inject Sheet18 code into the new sheet's code module
        if sheet18_code:
            try:
                new_sheet_vba_module = target_wb.VBProject.VBComponents(new_sheet.CodeName)
                new_sheet_vba_module.CodeModule.AddFromString(sheet18_code)
                print(f"Sheet18 VBA code injected successfully into {new_sheet.CodeName}.")
            except Exception as e:
                print(f"Error injecting Sheet18 VBA code: {e}")

        # Inject Module2 code into a new standard module
        if module2_code:
            try:
                new_module = target_wb.VBProject.VBComponents.Add(1) # Add standard module
                new_module.Name = "Module2"
                new_module.CodeModule.AddFromString(module2_code)
                print("Module2 VBA code injected successfully.")
            except Exception as e:
                print(f"Error injecting Module2 VBA code: {e}")

        target_wb.Save()

    except Exception as e:
        print(f"오류 발생: {e}")
    finally:
        if source_wb:
            source_wb.Close(False)
        if target_wb:
            target_wb.Close(False)
        if excel:
            excel.Quit()

if __name__ == "__main__":
    replicate_s_sheet()
