'==========================================================
' Excel VBA 매크로 - data.js 자동 생성
' 사용법:
'   1. Excel에서 Alt+F11 (VBA 편집기 열기)
'   2. 이 파일 내용을 모듈에 붙여넣기
'   3. ExportTasksData 매크로 실행 (Alt+F8)
'   또는 버튼에 연결해서 클릭으로 실행 가능
'
' [출력 상수 목록]
'   const ALL_DATA
'       과제 데이터  { cur_task, cmpt_task }
'       멤버 데이터  { bod, cnsl_partner, cnstc_partner }
'   const COLOR_MAPS
'       과제 색상맵  { task_keyward, comp_mngt, cmmp_exec }
'       멤버 색상맵  { role, career, cnsl_customer, cnst_keyward }
'==========================================================

Sub ExportTasksData()
    Dim basePath As String
    basePath = ThisWorkbook.Path & "\"

    Dim excelPath As String: excelPath = basePath & "data.xlsx"
    Dim outPath   As String: outPath   = basePath & "data.js"

    Dim excelWb As Workbook
    On Error GoTo ErrHandler

    ' 파일 열기 (표시 안 함)
    Application.ScreenUpdating = False
    Set excelWb = Workbooks.Open(excelPath, ReadOnly:=True)

    Dim js As String
    js = "// 자동 생성 파일 - Excel VBA 매크로로 갱신" & vbLf & _
         "// Excel 변경 후 '엑셀_데이터_내보내기' 매크로 재실행" & vbLf

    ' ── ① 색상 맵 (과제 + 멤버 병합 → COLOR_MAPS) ──────────
    Dim taskKwJson  As String: taskKwJson   = SheetToColorMap(excelWb.Sheets("task_keyward"),   "keyward", "color")
    Dim mngtJson    As String: mngtJson     = SheetToColorMap(excelWb.Sheets("comp_mngt"),      "comp",    "color")
    Dim execJson    As String
    On Error Resume Next
    execJson = SheetToColorMap(excelWb.Sheets("cmmp_exec"), "comp", "color")
    On Error GoTo ErrHandler
    If execJson = "" Then execJson = "{}"

    Dim roleJson     As String: roleJson     = SheetToColorMap(excelWb.Sheets("role"),          "role",     "color")
    Dim careerJson   As String: careerJson   = SheetToColorMap(excelWb.Sheets("career"),        "career",   "color")
    Dim cnslCustJson As String: cnslCustJson = SheetToColorMap(excelWb.Sheets("cnsl_customer"), "customer", "color")
    Dim cnstKwJson   As String: cnstKwJson   = SheetToColorMap(excelWb.Sheets("cnst_keyward"),  "keyward",  "color")

    Dim colorMaps As String
    colorMaps = "{" & vbLf & _
        "  ""task_keyward"":"  & taskKwJson   & "," & vbLf & _
        "  ""comp_mngt"":"     & mngtJson     & "," & vbLf & _
        "  ""cmmp_exec"":"     & execJson     & "," & vbLf & _
        "  ""role"":"          & roleJson     & "," & vbLf & _
        "  ""career"":"        & careerJson   & "," & vbLf & _
        "  ""cnsl_customer"":" & cnslCustJson & "," & vbLf & _
        "  ""cnst_keyward"":"  & cnstKwJson   & vbLf & _
        "}"

    ' ── ② 데이터 (과제 + 멤버 병합 → ALL_DATA) ─────────────
    Dim curJson   As String: curJson   = SheetToTaskArray(excelWb.Sheets("cur_task"))
    Dim cmptJson  As String: cmptJson  = SheetToTaskArray(excelWb.Sheets("cmpt_task"))
    Dim bodJson   As String: bodJson   = SheetToTaskArray(excelWb.Sheets("bod"))
    Dim cnslJson  As String: cnslJson  = SheetToTaskArray(excelWb.Sheets("cnsl_partner"))
    Dim cnstcJson As String: cnstcJson = SheetToTaskArray(excelWb.Sheets("cnstc_partner"))

    Dim allData As String
    allData = "{" & vbLf & _
        "  ""cur_task"":"      & curJson   & "," & vbLf & _
        "  ""cmpt_task"":"     & cmptJson  & "," & vbLf & _
        "  ""bod"":"           & bodJson   & "," & vbLf & _
        "  ""cnsl_partner"":"  & cnslJson  & "," & vbLf & _
        "  ""cnstc_partner"":" & cnstcJson & vbLf & _
        "}"

    js = js & "const ALL_DATA   = " & allData   & ";" & vbLf
    js = js & "const COLOR_MAPS = " & colorMaps & ";" & vbLf

    ' ── 파일 쓰기 (UTF-8 인코딩) ──────────────────────────
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Charset = "UTF-8"
    stream.Open
    stream.WriteText js
    stream.SaveToFile outPath, 2   ' 2 = 덮어쓰기
    stream.Close
    Set stream = Nothing

    ' 열었던 파일 닫기 (저장 안 함)
    excelWb.Close SaveChanges:=False

    Application.ScreenUpdating = True
    MsgBox "data.js 생성 완료!" & vbLf & outPath, vbInformation, "완료"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "오류: " & Err.Description, vbCritical, "오류"
End Sub


' ── 헬퍼: task 시트 → JSON 배열 ───────────────────────────
Private Function SheetToTaskArray(ws As Worksheet) As String
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' 헤더 읽기
    Dim headers() As String
    ReDim headers(1 To lastCol)
    Dim c As Long
    For c = 1 To lastCol
        headers(c) = Trim(CStr(ws.Cells(1, c).Value))
    Next c

    ' "title" → "name" 순서로 기준 컬럼 탐색 (없으면 첫 번째 컬럼)
    Dim titleIdx As Long: titleIdx = ColIdx(headers, "title")
    If titleIdx = 0 Then titleIdx = ColIdx(headers, "name")
    If titleIdx = 0 Then titleIdx = 1

    Dim items As String
    items = ""
    Dim r As Long
    For r = 2 To lastRow
        Dim titleVal As String: titleVal = Trim(CStr(ws.Cells(r, titleIdx).Value))
        If titleVal = "" Then GoTo NextRow

        Dim obj As String
        obj = "{"

        For c = 1 To lastCol
            Dim key As String: key = headers(c)
            If key = "" Then GoTo NextCol
            Dim cell As Range
            Set cell = ws.Cells(r, c)
            Dim val As String

            If key = "id" Then
                val = CStr(CLng(IIf(IsNumeric(cell.Value), cell.Value, r - 1)))
                obj = obj & """" & key & """:" & val & ","
            ElseIf key = "start_month" Or key = "end_month" Then
                Dim dv As Variant: dv = cell.Value
                If IsDate(dv) And Not IsEmpty(dv) Then
                    val = Format(CDate(dv), "yyyy-mm")
                ElseIf IsNumeric(dv) And Not IsEmpty(dv) Then
                    ' Excel 날짜 시리얼
                    val = Format(CDate(CLng(dv)), "yyyy-mm")
                Else
                    val = Trim(CStr(dv))
                End If
                obj = obj & """" & key & """:""" & JsonEsc(val) & ""","
            Else
                val = CStr(IIf(IsEmpty(cell.Value), "", cell.Value))
                obj = obj & """" & key & """:""" & JsonEsc(val) & ""","
            End If
NextCol:
        Next c

        ' 마지막 쉼표 제거
        If Right(obj, 1) = "," Then obj = Left(obj, Len(obj) - 1)
        obj = obj & "}"

        If items <> "" Then items = items & "," & vbLf & "    "
        items = items & obj
NextRow:
    Next r

    SheetToTaskArray = "[" & vbLf & "    " & items & vbLf & "  ]"
End Function


' ── 헬퍼: 색상 시트 → JSON 오브젝트 ─────────────────────────
Private Function SheetToColorMap(ws As Worksheet, keyCol As String, valCol As String) As String
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim headers() As String
    ReDim headers(1 To lastCol)
    Dim c As Long
    For c = 1 To lastCol
        headers(c) = Trim(CStr(ws.Cells(1, c).Value))
    Next c

    Dim ki As Long: ki = ColIdx(headers, keyCol)
    Dim vi As Long: vi = ColIdx(headers, valCol)
    If ki = 0 Or vi = 0 Then
        SheetToColorMap = "{}"
        Exit Function
    End If

    Dim pairs As String: pairs = ""
    Dim r As Long
    For r = 2 To lastRow
        Dim k As String: k = Trim(CStr(ws.Cells(r, ki).Value))
        Dim v As String: v = Trim(CStr(ws.Cells(r, vi).Value))
        If k <> "" And v <> "" Then
            If pairs <> "" Then pairs = pairs & ","
            pairs = pairs & """" & JsonEsc(k) & """:""" & JsonEsc(v) & """"
        End If
    Next r

    SheetToColorMap = "{" & pairs & "}"
End Function


' ── 헬퍼: 컬럼명 → 인덱스 ───────────────────────────────────
Private Function ColIdx(headers() As String, name As String) As Long
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        If headers(i) = name Then
            ColIdx = i
            Exit Function
        End If
    Next i
    ColIdx = 0
End Function


' ── 헬퍼: JSON 문자열 이스케이프 ────────────────────────────
Private Function JsonEsc(s As String) As String
    s = Replace(s, "\",  "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbTab, "\t")
    JsonEsc = s
End Function
