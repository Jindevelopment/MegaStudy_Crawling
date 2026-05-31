Sub QNA_분석_생성()

    Dim wsSource As Worksheet
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    
    Dim lastRow As Long
    Dim resultLastRow As Long
    Dim sumRow As Long
    
    Dim dict1 As Object
    Dim dict2 As Object
    
    Dim i As Long
    Dim targetDate As String
    Dim lectureName As String
    Dim categoryName As String
    
    Dim rowIndex As Long
    Dim isMatched As Boolean

    Set wsSource = Worksheets("QNA")

    Set dict1 = CreateObject("Scripting.Dictionary")
    Set dict2 = CreateObject("Scripting.Dictionary")

    ' =====================================================
    ' 기존 시트 삭제
    ' =====================================================

    Application.DisplayAlerts = False

    On Error Resume Next
    Worksheets("분석1").Delete
    Worksheets("분석2").Delete
    On Error GoTo 0

    Application.DisplayAlerts = True

    ' =====================================================
    ' 새 시트 생성
    ' =====================================================

    Set ws1 = Worksheets.Add
    ws1.Name = "분석1"

    Set ws2 = Worksheets.Add
    ws2.Name = "분석2"

    ' =====================================================
    ' [분석1] 머리글
    ' =====================================================

    ws1.Range("A1").Value = "날짜"
    ws1.Range("B1").Value = "한국지리"
    ws1.Range("C1").Value = "세계지리"
    ws1.Range("D1").Value = "동아시아사"
    ws1.Range("E1").Value = "세계사"
    ws1.Range("F1").Value = "생활과윤리"
    ws1.Range("G1").Value = "윤리와사상"
    ws1.Range("H1").Value = "사회문화"
    ws1.Range("I1").Value = "정치와법"
    ws1.Range("J1").Value = "경제"
    ws1.Range("K1").Value = "통합사회"
    ws1.Range("L1").Value = "한국사"
    ws1.Range("M1").Value = "그외"
    ws1.Range("N1").Value = "총 합계"

    ' =====================================================
    ' 원본 마지막 행
    ' =====================================================

    lastRow = wsSource.Cells(wsSource.Rows.Count, "E").End(xlUp).Row

    ' =====================================================
    ' 날짜 중복 제거
    ' =====================================================

    For i = 2 To lastRow

        If wsSource.Cells(i, "E").Value <> "" Then

            targetDate = Format(wsSource.Cells(i, "E").Value, "yyyy-mm-dd")

            If Not dict1.exists(targetDate) Then

                dict1.Add targetDate, dict1.Count + 2
                ws1.Cells(dict1(targetDate), "A").Value = targetDate

            End If

        End If

    Next i

    ' =====================================================
    ' 분석1 데이터 집계
    ' =====================================================

    For i = 2 To lastRow

        If wsSource.Cells(i, "E").Value <> "" Then

            targetDate = Format(wsSource.Cells(i, "E").Value, "yyyy-mm-dd")
            lectureName = Trim(wsSource.Cells(i, "D").Value)

            rowIndex = dict1(targetDate)

            isMatched = False

            If InStr(lectureName, "한국지리") > 0 Then
                ws1.Cells(rowIndex, "B").Value = ws1.Cells(rowIndex, "B").Value + 1
                isMatched = True
            End If

            If InStr(lectureName, "세계지리") > 0 Then
                ws1.Cells(rowIndex, "C").Value = ws1.Cells(rowIndex, "C").Value + 1
                isMatched = True
            End If

            If InStr(lectureName, "동아시아사") > 0 Then
                ws1.Cells(rowIndex, "D").Value = ws1.Cells(rowIndex, "D").Value + 1
                isMatched = True
            End If

            If InStr(lectureName, "세계사") > 0 Then
                ws1.Cells(rowIndex, "E").Value = ws1.Cells(rowIndex, "E").Value + 1
                isMatched = True
            End If

            If InStr(lectureName, "생활과윤리") > 0 Then
                ws1.Cells(rowIndex, "F").Value = ws1.Cells(rowIndex, "F").Value + 1
                isMatched = True
            End If

            If InStr(lectureName, "윤리와사상") > 0 Then
                ws1.Cells(rowIndex, "G").Value = ws1.Cells(rowIndex, "G").Value + 1
                isMatched = True
            End If

            If InStr(lectureName, "사회문화") > 0 Then
                ws1.Cells(rowIndex, "H").Value = ws1.Cells(rowIndex, "H").Value + 1
                isMatched = True
            End If

            If InStr(lectureName, "정치와법") > 0 Then
                ws1.Cells(rowIndex, "I").Value = ws1.Cells(rowIndex, "I").Value + 1
                isMatched = True
            End If

            If InStr(lectureName, "경제") > 0 Then
                ws1.Cells(rowIndex, "J").Value = ws1.Cells(rowIndex, "J").Value + 1
                isMatched = True
            End If

            If InStr(lectureName, "통합사회") > 0 Then
                ws1.Cells(rowIndex, "K").Value = ws1.Cells(rowIndex, "K").Value + 1
                isMatched = True
            End If

            If InStr(lectureName, "한국사") > 0 Then
                ws1.Cells(rowIndex, "L").Value = ws1.Cells(rowIndex, "L").Value + 1
                isMatched = True
            End If

            If isMatched = False Then
                ws1.Cells(rowIndex, "M").Value = ws1.Cells(rowIndex, "M").Value + 1
            End If

        End If

    Next i

    ' =====================================================
    ' 날짜 오름차순 정렬
    ' =====================================================

    With ws1.Sort

        .SortFields.Clear

        .SortFields.Add Key:=ws1.Range("A2:A" & ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal

        .SetRange ws1.Range("A1:N" & ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row)

        .Header = xlYes
        .Apply

    End With

    ' =====================================================
    ' 날짜별 총 합계
    ' =====================================================

    resultLastRow = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row

    For i = 2 To resultLastRow

        ws1.Cells(i, "N").Formula = "=SUM(B" & i & ":M" & i & ")"

    Next i

    ' =====================================================
    ' 합계 행
    ' =====================================================

    sumRow = resultLastRow + 1

    ws1.Cells(sumRow, "A").Value = "합계"

    For i = 2 To 14
        ws1.Cells(sumRow, i).Formula = "=SUM(" & _
        Split(ws1.Cells(1, i).Address, "$")(1) & "2:" & _
        Split(ws1.Cells(1, i).Address, "$")(1) & resultLastRow & ")"
    Next i

    ws1.Range("A" & sumRow & ":N" & sumRow).Font.Bold = True

    ws1.Columns("A:N").AutoFit

    ' =====================================================
    ' [분석2] 머리글
    ' =====================================================

    ws2.Range("A1").Value = "강좌명"
    ws2.Range("B1").Value = "강좌"
    ws2.Range("C1").Value = "교재"
    ws2.Range("D1").Value = "학습법"
    ws2.Range("E1").Value = "기타"
    ws2.Range("F1").Value = "총 합계"

    ' =====================================================
    ' 강좌명 중복 제거
    ' =====================================================

    For i = 2 To lastRow

        lectureName = Trim(wsSource.Cells(i, "D").Value)

        If lectureName <> "" Then

            If Not dict2.exists(lectureName) Then

                dict2.Add lectureName, dict2.Count + 2
                ws2.Cells(dict2(lectureName), "A").Value = lectureName

            End If

        End If

    Next i

    ' =====================================================
    ' 분석2 집계
    ' =====================================================

    For i = 2 To lastRow

        lectureName = Trim(wsSource.Cells(i, "D").Value)
        categoryName = Trim(wsSource.Cells(i, "C").Value)

        If lectureName <> "" Then

            rowIndex = dict2(lectureName)

            Select Case categoryName

                Case "[강좌]", "강좌"
                    ws2.Cells(rowIndex, "B").Value = ws2.Cells(rowIndex, "B").Value + 1

                Case "[교재]", "교재"
                    ws2.Cells(rowIndex, "C").Value = ws2.Cells(rowIndex, "C").Value + 1

                Case "[학습법]", "학습법"
                    ws2.Cells(rowIndex, "D").Value = ws2.Cells(rowIndex, "D").Value + 1

                Case "[기타]", "기타"
                    ws2.Cells(rowIndex, "E").Value = ws2.Cells(rowIndex, "E").Value + 1

            End Select

        End If

    Next i

    ' =====================================================
    ' 총 합계 계산
    ' =====================================================

    resultLastRow = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row

    For i = 2 To resultLastRow

        ws2.Cells(i, "F").Formula = "=SUM(B" & i & ":E" & i & ")"

    Next i

    ' =====================================================
    ' 합계 행
    ' =====================================================

    sumRow = resultLastRow + 1

    ws2.Cells(sumRow, "A").Value = "합계"

    ws2.Cells(sumRow, "B").Formula = "=SUM(B2:B" & resultLastRow & ")"
    ws2.Cells(sumRow, "C").Formula = "=SUM(C2:C" & resultLastRow & ")"
    ws2.Cells(sumRow, "D").Formula = "=SUM(D2:D" & resultLastRow & ")"
    ws2.Cells(sumRow, "E").Formula = "=SUM(E2:E" & resultLastRow & ")"
    ws2.Cells(sumRow, "F").Formula = "=SUM(F2:F" & resultLastRow & ")"

    ws2.Range("A" & sumRow & ":F" & sumRow).Font.Bold = True

    ws2.Columns("A:F").AutoFit

    MsgBox "분석1 / 분석2 시트 생성 완료!"

End Sub