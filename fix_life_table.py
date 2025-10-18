import openpyxl
from openpyxl.utils import get_column_letter

FILE_PATH = "C:\\InsuranceProject\\신한주니어큐브종합건강상해보험(무배당, 해약환급금 미지급형)_주보험.xlsx"

def get_input_data(workbook):
    """입력 시트에서 데이터를 읽어 딕셔너리로 반환합니다."""
    if "입력" not in workbook.sheetnames:
        raise ValueError("'입력' 시트를 찾을 수 없습니다.")

    sheet = workbook["입력"]
    data = {}
    # Find the cell for '예정이율'
    interest_rate_cell = None
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == "예정이율":
                interest_rate_cell = sheet.cell(row=cell.row, column=cell.column + 1).coordinate
                break
        if interest_rate_cell:
            break
    if not interest_rate_cell:
        # Assume B12 if not found
        interest_rate_cell = "B12"

    data['interest_rate_cell'] = interest_rate_cell

    for row in sheet.iter_rows(min_row=1, max_col=2):
        key_cell = row[0]
        value_cell = row[1]
        if key_cell.value:
            data[key_cell.value] = value_cell.value
    return data

def fix_life_table_with_formulas():
    """기수표를 동적 엑셀 수식으로 채웁니다."""
    try:
        workbook = openpyxl.load_workbook(FILE_PATH)

        # 1. 입력 데이터 읽기
        input_data = get_input_data(workbook)
        start_age = int(input_data['가입나이'])
        gender_cell = 'B4' # Assuming gender is in B4 of '입력' sheet
        interest_rate_cell = input_data['interest_rate_cell']
        max_age = 120 # Maximum age for the table

        # 2. 기수표 시트 준비
        if "기수표" not in workbook.sheetnames:
            workbook.create_sheet("기수표")
        sheet = workbook["기수표"]
        sheet.delete_rows(1, sheet.max_row)  # 기존 데이터 삭제

        # Header 쓰기
        header = ['연령(x)', '생존자수(lx)', '사망률(qx)', '사망자수(dx)', 'Dx', 'Nx', 'Cx', 'Mx']
        sheet.append(header)

        # 3. 수식 작성
        for age in range(start_age, max_age + 1):
            row = age - start_age + 2
            sheet.cell(row=row, column=1, value=age)

            # lx (생존자수)
            if age == start_age:
                sheet.cell(row=row, column=2, value=100000)
            else:
                sheet.cell(row=row, column=2, value=f'=B{row}-D{row}')

            # qx (사망률)
            # VLOOKUP(age, risk_table_range, IF(gender="남", 2, 3), FALSE)
            gender_condition = f'IF(입력!${gender_cell}="남", 2, 3)'
            # Assuming the risk table is in '위험률' sheet from A3 to C103
            risk_table_range = '위험률!$A$3:$C$103' 
            sheet.cell(row=row, column=3, value=f'=VLOOKUP(A{row}, {risk_table_range}, {gender_condition}, FALSE)')

            # dx (사망자수)
            sheet.cell(row=row, column=4, value=f'=B{row}*C{row}')

            # Dx
            v_formula = f'(1/(1+입력!${interest_rate_cell}))'
            sheet.cell(row=row, column=5, value=f'=B{row}*{v_formula}^A{row}')

            # Cx
            sheet.cell(row=row, column=7, value=f'=D{row}*{v_formula}^(A{row}+1)')

        # Nx, Mx (역산)
        max_row = max_age - start_age + 2
        for age in range(start_age, max_age + 1):
            row = age - start_age + 2
            # Nx
            sheet.cell(row=row, column=6, value=f'=SUM(E{row}:E${max_row})')
            # Mx
            sheet.cell(row=row, column=8, value=f'=SUM(G{row}:G${max_row})')

        # 5. 파일 저장
        workbook.save(FILE_PATH)
        print("기수표가 동적 수식으로 성공적으로 수정되었습니다.")

    except Exception as e:
        print(f"오류 발생: {e}")

if __name__ == "__main__":
    fix_life_table_with_formulas()
