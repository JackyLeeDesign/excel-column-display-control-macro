
Private Sub Worksheet_Activate()
    CheckAllCellNames
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    CheckAllCellNames
End Sub

Function CheckAllCellNames()
    Dim targetRange As Range
    Dim targetColumn As String
    Dim targetValue As String
    Dim targetRangeStr As String
    Dim cell As Range
    Dim conditions() As String
    Dim condition As Variant
    Dim rangeStr As String
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim nm As Name
    Dim i As Integer

    '獲取當前工作簿
    Set wb = ActiveWorkbook
    '獲取當前工作表
    Set ws = wb.ActiveSheet
    '獲取當前工作表的所有命名區域s
    For Each nm In wb.Names
        If InStr(1, nm.RefersTo, ws.Name) > 0 Then
            ShowOrHideRows nm.Name, nm.RefersTo
        End If
    Next nm
End Function

Function ShowOrHideRows(fieldName As String, relatedRange As String)
    Dim fieldNameParts() As String
    fieldNameParts = Split(fieldName, "_")
    '全形轉半形
    fieldNameParts(0) = StrConv(fieldNameParts(0), vbNarrow)
    '小寫轉大寫
    fieldNameParts(0) = UCase(fieldNameParts(0))

    Dim columnName As String
    columnName = fieldNameParts(0)
    '全形轉半形
    columnName = StrConv(columnName, vbNarrow)
    '小寫轉大寫
    columnName = UCase(columnName)

    Dim fieldValue As String
    fieldValue = fieldNameParts(1)
    '全形轉半形
    fieldValue = StrConv(fieldValue, vbNarrow)
    '小寫轉大寫
    fieldValue = UCase(fieldValue)

    Dim actionValue As String
    actionValue = fieldNameParts(2)
    '全形轉半形
    actionValue = StrConv(actionValue, vbNarrow)
    '小寫轉大寫
    actionValue = UCase(actionValue)

    'relatedRange=Sheet1!$A$3:$B$4，擷取Sheet1 存入 sheetName
    Dim sheetName As String
    sheetName = Split(relatedRange, "!")(0)

    '擷取命名中之起始欄位範圍，存入 startCell，若僅為"=Sheet1!$7:$8"，則為 A7
    Dim startCell As String
    If InStr(1, Split(relatedRange, "!")(1), ":") > 0 Then
        startCell = Split(Split(relatedRange, "!")(1), ":")(0)
    Else
        startCell = Split(relatedRange, "!")(1)
    End If

    '擷取命名中之結束欄位範圍，存入 endCell，若僅為"=Sheet1!$7:$8"，則為 A7，若無則為 startCell
    Dim endCell As String
    If InStr(1, Split(relatedRange, "!")(1), ":") > 0 Then
        endCell = Split(Split(relatedRange, "!")(1), ":")(1)
    Else
        endCell = startCell
    End If
    
    Dim targetRange As Range
    Set targetRange = Range(sheetName & "!" & startCell & ":" & endCell)
    '獲取目標欄位的值
    targetValue = Range(columnName).Value
    '全形轉半形
    targetValue = StrConv(targetValue, vbNarrow)
    '小寫轉大寫
    targetValue = UCase(targetValue)

    If targetValue = fieldValue Then
        If actionValue = "SHOW" Then
            targetRange.EntireRow.Hidden = False
        ElseIf actionValue = "HIDE" Then
            targetRange.EntireRow.Hidden = True
        End If
    ElseIf targetValue <> fieldValue Then
        If actionValue = "SHOW" Then
            targetRange.EntireRow.Hidden = True
        ElseIf actionValue = "HIDE" Then
            targetRange.EntireRow.Hidden = False
        End If
    End If
End Function