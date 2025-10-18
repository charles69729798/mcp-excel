
Sub CreateInsuranceSheet()
    Dim ws As Worksheet
    Dim mappingWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim mappingCounter As Long
    Dim startMapRow As Long
    Dim endMapRow As Long
    Dim dropdownValues As String
    Dim formulaString As String
    
    Application.DisplayAlerts = False
    On Error Resume Next
    Worksheets("입력").Delete
    Worksheets("data_mapping").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "입력"
    Set mappingWs = ThisWorkbook.Sheets.Add(After:=ws)
    mappingWs.Name = "data_mapping"
    mappingWs.Visible = xlSheetHidden
    
    mappingCounter = 1
    
    ws.Cells(1, 1).Value = "구분"
    ws.Cells(1, 2).Value = "항목"
    ws.Cells(1, 3).Value = "입력값"
    ws.Cells(1, 4).Value = "비고 / 선택 옵션"
    ws.Cells(1, 5).Value = "관련 근거"
    ws.Cells(1, 6).Value = "참조 값"
    
    Dim jsonData As Variant
    jsonData = Array( _
        Array("공통", "가입 유형", "드롭다운", "어린이형,태아형", True, "2.가, 2.나"), _
        Array("공통", "상품 종류", "드롭다운", "해약환급금 미지급형,일반형", True, "보험종목의 명칭, 2.가"), _
        Array("주계약", "갱신 여부", "자동 표시", "비갱신형", False, "보험종목의 명칭"), _
        Array("주계약", "보험 기간", "드롭다운", "30세만기,100세만기", True, "2.가"), _
        Array("주계약", "납입 기간", "드롭다운", "5년납,10년납,20년납,30년납,전기납", True, "2.가"), _
        Array("일반암진단특약", "갱신 여부", "자동 표시", "비갱신형", False, "보험종목의 명칭"), _
        Array("일반암진단특약", "보험 기간", "드롭다운", "30세만기,100세만기", True, "2.가"), _
        Array("일반암진단특약", "납입 기간", "드롭다운", "5년납,10년납,20년납,30년납,전기납", True, "2.가"), _
        Array("고액암진단특약", "갱신 여부", "자동 표시", "비갱신형", False, "보험종목의 명칭"), _
        Array("고액암진단특약", "보험 기간", "드롭다운", "30세만기,100세만기", True, "2.가"), _
        Array("고액암진단특약", "납입 기간", "드롭다운", "5년납,10년납,20년납,30년납,전기납", True, "2.가"), _
        Array("남녀특정암진단특약", "갱신 여부", "자동 표시", "비갱신형", False, "보험종목의 명칭"), _
        Array("남녀특정암진단특약", "보험 기간", "드롭다운", "30세만기,100세만기", True, "2.가"), _
        Array("남녀특정암진단특약", "납입 기간", "드롭다운", "5년납,10년납,20년납,30년납,전기납", True, "2.가"), _
        Array("소액암진단특약", "갱신 여부", "자동 표시", "비갱신형", False, "보험종목의 명칭"), _
        Array("소액암진단특약", "보험 기간", "드롭다운", "30세만기,100세만기", True, "2.가"), _
        Array("소액암진단특약", "납입 기간", "드롭다운", "5년납,10년납,20년납,30년납,전기납", True, "2.가"), _
        Array("주니어특정중대질병진단특약", "갱신 여부", "자동 표시", "비갱신형", False, "보험종목의 명칭"), _
        Array("주니어특정중대질병진단특약", "보험 기간", "드롭다운", "30세만기,100세만기", True, "2.가"), _
        Array("주니어특정중대질병진단특약", "납입 기간", "드롭다운", "5년납,10년납,20년납,30년납,전기납", True, "2.가"), _
        Array("주니어특정중대질병수술보장특약", "갱신 여부", "자동 표시", "비갱신형", False, "보험종목의 명칭"), _
        Array("주니어특정중대질병수술보장특약", "보험 기간", "드롭다운", "30세만기,100세만기", True, "2.가"), _
        Array("주니어특정중대질병수술보장특약", "납입 기간", "드롭다운", "5년납,10년납,20년납,30년납,전기납", True, "2.가"), _
        Array("뇌출혈진단특약", "갱신 여부", "자동 표시", "비갱신형", False, "보험종목의 명칭"), _
        Array("뇌출혈진단특약", "보험 기간", "드롭다운", "100세만기", True, "2.가"), _
        Array("뇌출혈진단특약", "납입 기간", "드롭다운", "10년납,20년납,30년납", True, "2.가"), _
        Array("급성심근경색증진단특약", "갱신 여부", "자동 표시", "비갱신형", False, "보험종목의 명칭"), _
        Array("급성심근경색증진단특약", "보험 기간", "드롭다운", "100세만기", True, "2.가"), _
        Array("급성심근경색증진단특약", "납입 기간", "드롭다운", "10년납,20년납,30년납", True, "2.가"), _
        Array("뇌혈관질환진단특약", "갱신 여부", "자동 표시", "비갱신형", False, "보험종목의 명칭"), _
        Array("뇌혈관질환진단특약", "보험 기간", "드롭다운", "100세만기", True, "2.가"), _
        Array("뇌혈관질환진단특약", "납입 기간", "드롭다운", "10년납,20년납,30년납", True, "2.가"), _
        Array("허혈심장질환진단특약", "갱신 여부", "자동 표시", "비갱신형", False, "보험종목의 명칭"), _
        Array("허혈심장질환진단특약", "보험 기간", "드롭다운", "100세만기", True, "2.가"), _
        Array("허혈심장질환진단특약", "납입 기간", "드롭다운", "10년납,20년납,30년납", True, "2.가"), _
        Array("질병장해특약", "갱신 여부", "자동 표시", "비갱신형", False, "보험종목의 명칭"), _
        Array("질병장해특약", "보험 기간", "드롭다운", "30세만기,100세만기", True, "2.가"), _
        Array("질병장해특약", "납입 기간", "드롭다운", "5년납,10년납,20년납,30년납,전기납", True, "2.가"), _
        Array("첫날부터플러스입원보장특약", "갱신 여부", "자동 표시", "비갱신형", False, "보험종목의 명칭"), _
        Array("첫날부터플러스입원보장특약", "보험 기간", "드롭다운", "30세만기,100세만기", True, "2.가"), _
        Array("첫날부터플러스입원보장특약", "납입 기간", "드롭다운", "5년납,10년납,20년납,30년납,전기납", True, "2.가"), _
        Array("응급실내원특약", "갱신 여부", "자동 표시", "비갱신형", False, "보험종목의 명칭"), _
        Array("응급실내원특약", "보험 기간", "드롭다운", "30세만기,100세만기", True, "2.가"), _
        Array("응급실내원특약", "납입 기간", "드롭다운", "5년납,10년납,20년납,30년납,전기납", True, "2.가"), _
        Array("수술특약", "갱신 여부", "자동 표시", "비갱신형", False, "보험종목의 명칭"), _
        Array("수술특약", "보험 기간", "드롭다운", "30세만기,100세만기", True, "2.가"), _
        Array("수술특약", "납입 기간", "드롭다운", "5년납,10년납,20년납,30년납,전기납", True, "2.가"), _
        Array("수술플러스보장특약", "갱신 여부", "자동 표시", "비갱신형", False, "보험종목의 명칭"), _
        Array("수술플러스보장특약", "보험 기간", "드롭다운", "30세만기,100세만기", True, "2.가"), _
        Array("수술플러스보장특약", "납입 기간", "드롭다운", "5년납,10년납,20년납,30년납,전기납", True, "2.가"), _
        Array("6대질병수술보장특약", "갱신 여부", "자동 표시", "비갱신형", False, "보험종목의 명칭"), _
        Array("6대질병수술보장특약", "보험 기간", "드롭다운", "30세만기,100세만기", True, "2.가"), _
        Array("6대질병수술보장특약", "납입 기간", "드롭다운", "5년납,10년납,20년납,30년납,전기납", True, "2.가"), _
        Array("추간판장애및컴퓨터과잉질환수술특약", "갱신 여부", "자동 표시", "비갱신형", False, "보험종목의 명칭"), _
        Array("추간판장애및컴퓨터과잉질환수술특약", "보험 기간", "드롭다운", "30세만기,100세만기", True, "2.가"), _
        Array("추간판장애및컴퓨터과잉질환수술특약", "납입 기간", "드롭다운", "5년납,10년납,20년납,30년납,전기납", True, "2.가"), _
        Array("골절깁스치료특약", "갱신 여부", "자동 표시", "비갱신형", False, "보험종목의 명칭"), _
        Array("골절깁스치료특약", "보험 기간", "드롭다운", "30세만기,100세만기", True, "2.가"), _
        Array("골절깁스치료특약", "납입 기간", "드롭다운", "5년납,10년납,20년납,30년납,전기납", True, "2.가"), _
        Array("주니어특정질병입원특약", "갱신 여부", "자동 표시", "비갱신형", False, "보험종목의 명칭"), _
        Array("주니어특정질병입원특약", "보험 기간", "드롭다운", "30세만기", True, "2.가"), _
        Array("주니어특정질병입원특약", "납입 기간", "드롭다운", "5년납,10년납,전기납", True, "2.가"), _
        Array("피부질환수술특약", "갱신 여부", "자동 표시", "비갱신형", False, "보험종목의 명칭"), _
        Array("피부질환수술특약", "보험 기간", "드롭다운", "30세만기,100세만기", True, "2.가"), _
        Array("피부질환수술특약", "납입 기간", "드롭다운", "5년납,10년납,20년납,30년납,전기납", True, "2.가"), _
        Array("수술특약(갱신형)", "갱신 여부", "자동 표시", "갱신형", False, "보험종목의 명칭"), _
        Array("수술특약(갱신형)", "보험 기간", "드롭다운", "5년", True, "2.가"), _
        Array("수술특약(갱신형)", "납입 기간", "드롭다운", "전기납", True, "2.가"), _
        Array("첫날부터입원보장특약(갱신형)", "갱신 여부", "자동 표시", "갱신형", False, "보험종목의 명칭"), _
        Array("첫날부터입원보장특약(갱신형)", "보험 기간", "드롭다운", "5년", True, "2.가"), _
        Array("첫날부터입원보장특약(갱신형)", "납입 기간", "드롭다운", "전기납", True, "2.가"), _
        Array("첫날부터플러스입원보장특약(갱신형)", "갱신 여부", "자동 표시", "갱신형", False, "보험종목의 명칭"), _
        Array("첫날부터플러스입원보장특약(갱신형)", "보험 기간", "드롭다운", "5년", True, "2.가"), _
        Array("첫날부터플러스입원보장특약(갱신형)", "납입 기간", "드롭다운", "전기납", True, "2.가") _
    )
    
    For i = LBound(jsonData) To UBound(jsonData)
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        
        ws.Cells(lastRow, 1).Value = jsonData(i)(0)
        ws.Cells(lastRow, 2).Value = jsonData(i)(1)
        ws.Cells(lastRow, 5).Value = jsonData(i)(5)
        
        If jsonData(i)(2) = "드롭다운" Then
            ws.Cells(lastRow, 4).Value = "선택: " & jsonData(i)(3)
            
            With ws.Cells(lastRow, 3).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=jsonData(i)(3)
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            
            If jsonData(i)(4) = True Then
                startMapRow = mappingCounter
                Dim mappingKeys As Variant
                mappingKeys = Split(jsonData(i)(3), ",")
                
                Dim j As Long
                For j = LBound(mappingKeys) To UBound(mappingKeys)
                    mappingWs.Cells(mappingCounter, 1).Value = mappingKeys(j)
                    mappingWs.Cells(mappingCounter, 2).Value = Val(mappingKeys(j))
                    mappingCounter = mappingCounter + 1
                Next j
                endMapRow = mappingCounter - 1
                
                formulaString = "=IFERROR(VLOOKUP(C" & lastRow & ",data_mapping!A" & startMapRow & ":B" & endMapRow & ",2,FALSE), """")"
                ws.Cells(lastRow, 6).Formula = formulaString
            End If
            
        ElseIf jsonData(i)(2) = "숫자 입력" Then
            ws.Cells(lastRow, 4).Value = "사용자 직접 입력"
        Else
            ws.Cells(lastRow, 3).Value = jsonData(i)(3)
        End If
    Next i
    
    ws.Columns.AutoFit
    
    MsgBox "시트 생성이 완료되었습니다."

End Sub
