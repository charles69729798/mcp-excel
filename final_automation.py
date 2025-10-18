import os
import json
import win32com.client as win32
import google.generativeai as genai
from pypdf import PdfReader

# --- 1. 설정 ---
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "AIzaSyCNiN_y_8DAyuX9iFWgG6HK_mlEXlmsspQ")
PDF_FILE_PATH = "C:\\InsuranceProject\\034_사업방법서_신한주니어큐브종합건강상해보험(무배당, 해약환급금 미지급형)_20240401.pdf"

# --- 2. PDF 및 API 관련 함수 ---
def extract_text_from_pdf(pdf_path):
    print("1/5: PDF 파일에서 텍스트를 추출하는 중...")
    try:
        reader = PdfReader(pdf_path)
        text = ""
        for page in reader.pages:
            text += page.extract_text() or ""
        return text
    except Exception as e:
        print(f"PDF 처리 중 오류: {e}")
        return None

def get_analysis_prompt(document_text):
    return f'''
    당신은 보험 상품 분석 전문가입니다. 주어진 '사업방법서' 텍스트를 분석하여, 사용자가 엑셀에서 보험료 계산을 시뮬레이션하는 데 필요한 입력 항목들을 추출하고, 이를 JSON 형식으로 구조화해주세요.

    **분석 규칙:**
    1. '출력값 매핑': '입력 방식'이 '드롭다운'인 항목에 대해, 각 '선택 가능 값'이 어떤 숫자/참조 값에 해당하는지 매핑 정보를 생성합니다. (예: '10년납'은 10, '30세만기'는 30). 이 매핑 정보가 없으면 null로 설정합니다.

    **출력 형식 (JSON):**
    - 반드시 다음 키를 포함하는 JSON 객체의 배열(list of objects)로 응답해주세요.
    - `["구분", "항목", "입력 방식", "선택 가능 값", "출력값 매핑", "관련 근거"]`
    - "출력값 매핑"은 JSON 객체(object) 형식이어야 합니다.

    **예시 JSON 객체:**
    {{
        "구분": "주계약",
        "항목": "납입 기간",
        "입력 방식": "드롭다운",
        "선택 가능 값": "10년납,20년납,30년납",
        "출력값 매핑": {{
            "10년납": 10,
            "20년납": 20,
            "30년납": 30
        }},
        "관련 근거": "사업방법서 2.가"
    }}

    이제 아래의 사업방법서 텍스트를 분석하여 완전한 JSON 배열을 생성해주세요.

    --- 문서 텍스트 시작 ---
    {document_text}
    --- 문서 텍스트 끝 ---
    '''

def analyze_document_with_gemini(api_key, document_text):
    print("2/5: Gemini API를 호출하여 문서를 분석하는 중...")
    if not api_key or api_key == "YOUR_API_KEY_HERE":
        print("오류: Gemini API 키가 설정되지 않았습니다.")
        return None
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('models/gemini-2.5-pro')
        prompt = get_analysis_prompt(document_text)
        response = model.generate_content(prompt)
        
        json_start_index = response.text.find('[')
        json_end_index = response.text.rfind(']')
        if json_start_index == -1 or json_end_index == -1:
            print("오류: API 응답에서 유효한 JSON을 찾을 수 없습니다.")
            return None
        json_string = response.text[json_start_index : json_end_index + 1]
        return json.loads(json_string)
    except Exception as e:
        print(f"API 호출 중 오류: {e}")
        return None

# --- 3. VBA 코드 생성 함수 ---
def generate_vba_code(json_data):
    print("3/5: 분석된 데이터를 바탕으로 VBA 코드를 동적으로 생성하는 중...")
    
    vba_lines = []
    vba_lines.append("Sub CreateInsuranceSheet()")
    vba_lines.append("    Dim ws As Worksheet, mappingWs As Worksheet")
    vba_lines.append("    Application.DisplayAlerts = False")
    vba_lines.append("    On Error Resume Next")
    vba_lines.append("    Worksheets(\"입력\").Delete")
    vba_lines.append("    Worksheets(\"data_mapping\").Delete")
    vba_lines.append("    On Error GoTo 0")
    vba_lines.append("    Application.DisplayAlerts = True")
    vba_lines.append("    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))")
    vba_lines.append("    ws.Name = \"입력\"")
    vba_lines.append("    Set mappingWs = ThisWorkbook.Sheets.Add(After:=ws)")
    vba_lines.append("    mappingWs.Name = \"data_mapping\"")
    vba_lines.append("    mappingWs.Visible = xlSheetHidden")
    vba_lines.append("    ws.Range(\"A1:F1\").Value = Array(\"구분\", \"항목\", \"입력값\", \"비고 / 선택 옵션\", \"관련 근거\", \"참조 값\")")

    # Python 변수로 매핑 시트의 행 카운터 초기화
    mapping_counter = 1
    row_idx = 2

    for item in json_data:
        vba_lines.append(f'    ws.Cells({row_idx}, 1).Value = "{item.get("구분", "")}"')
        vba_lines.append(f'    ws.Cells({row_idx}, 2).Value = "{item.get("항목", "")}"')
        vba_lines.append(f'    ws.Cells({row_idx}, 5).Value = "{item.get("관련 근거", "")}"')

        input_type = item.get("입력 방식", "")
        if input_type == "드롭다운":
            dropdown_values = item.get("선택 가능 값", "")
            vba_lines.append(f'    ws.Cells({row_idx}, 4).Value = "선택: {dropdown_values}"')
            vba_lines.append(f'    With ws.Cells({row_idx}, 3).Validation')
            vba_lines.append("        .Delete")
            vba_lines.append(f'        .Add Type:=xlValidateList, Formula1:="{dropdown_values}"')
            vba_lines.append("    End With")
            
            mapping_info = item.get("출력값 매핑")
            if mapping_info:
                start_map_row = mapping_counter
                for key, value in mapping_info.items():
                    vba_lines.append(f'    mappingWs.Cells({mapping_counter}, 1).Value = "{key}"')
                    vba_lines.append(f'    mappingWs.Cells({mapping_counter}, 2).Value = "{value}"')
                    mapping_counter += 1
                end_map_row = mapping_counter - 1
                vba_lines.append(f'    ws.Cells({row_idx}, 6).Formula = "=IFERROR(VLOOKUP(C{row_idx}, data_mapping!A{start_map_row}:B{end_map_row}, 2, FALSE), \"\")"')

        elif input_type == "숫자 입력":
            vba_lines.append(f'    ws.Cells({row_idx}, 4).Value = "사용자 직접 입력"')
        elif input_type == "자동 표시":
            vba_lines.append(f'    ws.Cells({row_idx}, 3).Value = "{item.get("선택 가능 값", "")}"')
        
        row_idx += 1

    vba_lines.append("    ws.Columns.AutoFit")
    vba_lines.append("    MsgBox \"시트 생성이 완료되었습니다.\"")
    vba_lines.append("End Sub")
    
    return "\n".join(vba_lines)

# --- 4. 엑셀 제어 및 저장 함수 ---
def inject_vba_and_save(vba_code, output_path):
    print("4/5: 엑셀을 실행하고 매크로를 주입하는 중...")
    excel = None
    try:
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False
        workbook = excel.Workbooks.Add()
        vba_module = workbook.VBProject.VBComponents.Add(1)
        vba_module.CodeModule.AddFromString(vba_code)
        workbook.SaveAs(output_path, FileFormat=52)
    except Exception as e:
        print(f"엑셀 제어 중 오류: {e}")
        print("엑셀이 설치되어 있는지, 보안 설정에서 'VBA 프로젝트 개체 모델에 대한 액세스 신뢰'가 활성화되어 있는지 확인하세요.")
    finally:
        if excel:
            workbook.Close(SaveChanges=False)
            excel.Quit()

# --- 5. 메인 실행 블록 ---
if __name__ == "__main__":
    # API 키 확인
    if not GEMINI_API_KEY or GEMINI_API_KEY == "YOUR_API_KEY_HERE":
        print("오류: 스크립트 상단의 GEMINI_API_KEY를 본인의 키로 수정해주세요.")
    else:
        # 1. PDF 분석
        doc_text = extract_text_from_pdf(PDF_FILE_PATH)
        if doc_text:
            # 2. API 호출로 JSON 데이터 확보
            json_data = analyze_document_with_gemini(GEMINI_API_KEY, doc_text)
            if json_data:
                # 3. VBA 코드 생성
                vba_code = generate_vba_code(json_data)
                if vba_code:
                    # 4. 엑셀에 주입 및 저장
                    final_path = os.path.join(os.getcwd(), "최종_보험료산출.xlsm")
                    inject_vba_and_save(vba_code, final_path)
                    print(f"\n5/5: 작업 완료! '{final_path}' 파일이 생성되었습니다.")
                    print("파일을 열고 '콘텐츠 사용' 버튼을 누른 후, Alt+F8로 'CreateInsuranceSheet' 매크로를 실행하세요.")