' 將底下 Worksheet_Activate 和  Worksheet_Change 放在需要執行巨集的工作表內。
Private Sub Worksheet_Activate()
    CheckAllCellNames
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    CheckAllCellNames
End Sub

' 建立一個模組，將底下程式碼放在模組內。
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
    '獲取當前工作表的所有命名區域
    For Each nm In wb.Names
        If InStr(1, nm.RefersTo, ws.Name) > 0 Then
            ShowOrHideRows nm.Name, nm.RefersTo
        End If
    Next nm
End Function
Function ShowOrHideRows(fieldName As String, relatedRange As String)
    'ex: "B2.YES_and_B3.NO_or_B4.YES__SHOW"
    '將條件分割
    Dim conditionsStr As String
    conditionsStr = Split(fieldName, "__")(0)

    Dim actionValue As String
    actionValue = Split(fieldName, "__")(1)
    '全形轉半形
    actionValue = StrConv(actionValue, vbNarrow)
    '小寫轉大寫
    actionValue = UCase(actionValue)
    
    '獲取sheetName
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

    '檢查每個條件, 若有一個條件不符合, 則不顯示
    If CheckCondition(conditionsStr) Then
        If actionValue = "SHOW" Then
            targetRange.EntireRow.Hidden = False
        ElseIf actionValue = "HIDE" Then
            targetRange.EntireRow.Hidden = True
        End If
    ElseIf Not CheckCondition(conditionsStr) Then
        If actionValue = "SHOW" Then
            targetRange.EntireRow.Hidden = True
        ElseIf actionValue = "HIDE" Then
            targetRange.EntireRow.Hidden = False
        End If
    End If
End Function
Function CheckCondition(condition As String) As Boolean
    '宣告 ResultCondition 為字串
    Dim ResultCondition As String
    ResultCondition = condition

    '宣告 ResultBool 為布林值
    Dim ResultBool As Boolean
    
    Dim inputTmpString As String

    '將字串內出現的 "_and_" 與 "_or_" 替換成 "|"
    inputTmpString = Replace(condition, "_and_", "|")
    inputTmpString = Replace(condition, "_or_", "|")

    '宣告條件陣列
    Dim columnInfoArray() As String
    '將字串以 "|" 分割成陣列
    columnInfoArray = Split(inputTmpString, "|")

    '逐筆檢查條件
    Dim columnInfo As Variant
    For Each columnInfo In columnInfoArray
        ResultCondition = Replace(ResultCondition, columnInfo, CheckFieldValue(columnInfo))
    Next columnInfo

    '計算 ResultCondition 的結果
    ResultBool = Evaluate(ResultCondition)
    CheckCondition = ResultBool
End Function
Function CheckFieldValue(columnInfo As Variant) As String
    'B2.YES
    '將字串以 "." 分割成陣列
    Dim fieldNameParts() As String
    fieldNameParts = Split(columnInfo, ".")

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

    '獲取目標欄位的值
    targetValue = Range(columnName).Value
    '全形轉半形
    targetValue = StrConv(targetValue, vbNarrow)
    '小寫轉大寫
    targetValue = UCase(targetValue)

    If targetValue = fieldValue Then
        CheckFieldValue = "True"
    ElseIf targetValue <> fieldValue Then
        CheckFieldValue = "False"
    End If
End Function
