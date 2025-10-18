
import win32com.client
import os

FILE_PATH = os.path.abspath("C:\\InsuranceProject\\신한주니어큐브종합건강상해보험(무배당, 해약환급금 미지급형)_주보험.xlsx")

def fix_life_table_with_formulas_win32():
    """기수표를 동적 엑셀 수식으로 채우고 win32com을 사용하여 재계산합니다."""
    excel = None
    workbook = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        workbook = excel.Workbooks.Open(FILE_PATH)

        # 1. 입력 데이터 읽기
        input_sheet = workbook.Sheets("입력")
        start_age = int(input_sheet.Range("B3").Value)
        gender = input_sheet.Range("B4").Value
        # Find interest rate cell
        interest_rate_cell_address = None
        for cell in input_sheet.UsedRange:
            if cell.Value == "예정이율":
                interest_rate_cell_address = cell.Offset(0, 2).Address
                break
        if not interest_rate_cell_address:
            interest_rate_cell_address = "$D$11"

        # 2. 기수표 시트 준비
        try:
            sheet = workbook.Sheets("기수표")
        except:
            sheet = workbook.Sheets.Add(After=workbook.Sheets(workbook.Sheets.Count))
            sheet.Name = "기수표"
        sheet.Cells.Clear()

        # Header 쓰기
        header = ['연령(x)', '생존자수(lx)', '사망률(qx)', '사망자수(dx)', 'Dx', 'Nx', 'Cx', 'Mx']
        for i, h in enumerate(header):
            sheet.Cells(1, i + 1).Value = h

        # 3. 수식 작성
        max_age = 120
        for age in range(start_age, max_age + 1):
            row = age - start_age + 2
            sheet.Cells(row, 1).Value = age

            if age == start_age:
                sheet.Cells(row, 2).Value = 100000
            else:
                sheet.Cells(row, 2).Formula = f'=B{row - 1}-D{row - 1}'

            gender_condition = f'IF(입력!$B$4="남", 2, 3)'
            risk_table_range = '위험률_확장!$A$3:$C$123'
            sheet.Cells(row, 3).Formula = f'=VLOOKUP(A{row}, {risk_table_range}, {gender_condition}, FALSE)'

            sheet.Cells(row, 4).Formula = f'=B{row}*C{row}'

            v_formula = '(1/(1+0.0225))'
            sheet.Cells(row, 5).Formula = f'=B{row}*{v_formula}^A{row}'

            sheet.Cells(row, 7).Formula = f'=D{row}*{v_formula}^(A{row}+1)'

        max_row = max_age - start_age + 2
        for age in range(start_age, max_age + 1):
            row = age - start_age + 2
            sheet.Cells(row, 6).Formula = f'=SUM(E{row}:E${max_row})'
            sheet.Cells(row, 8).Formula = f'=SUM(G{row}:G${max_row})'

        # 4. 재계산
        workbook.Application.Calculate()

        # 5. 파일 저장
        workbook.Save()
        print("기수표가 동적 수식으로 성공적으로 수정되고 재계산되었습니다.")

    except Exception as e:
        print(f"오류 발생: {e}")
    finally:
        if workbook:
            workbook.Close(False)
        if excel:
            excel.Quit()

if __name__ == "__main__":
    fix_life_table_with_formulas_win32()
