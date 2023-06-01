' 將底下 Worksheet_Activate 和  Worksheet_Change 放在需要執行巨集的工作表內。
Private Sub Worksheet_Activate()
    CheckAllCellNames
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    CheckAllCellNames
End Sub

' 建立一個模組，將底下程式碼放在模組內。
Function CheckAllCellNames()
    On Error GoTo ErrorHandler
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
        For Each nm In ws.Names
            If InStr(1, nm.Name, "__") > 0 And InStr(1, nm.Name, ".") > 0 Then
                isDo = True
                '若命名區域參照範圍包含多個範圍，則跳出提醒視窗，顯示: nm.Name & " 命名區域之參照範圍包含多個範圍，請重新確認。"
                If InStr(1, nm.RefersTo, "+") > 0 Then
                    MsgBox nm.Name & " 命名區域之參照範圍包含多個範圍，請重新確認。"
                    isDo = False
                End If
                If isDo = True Then
                    ShowOrHideRows nm.Name, nm.RefersTo
                End If
            End If
        Next nm
    Exit Function
    ErrorHandler:
        MsgBox "程式在讀取命名規則時發生錯誤，規則名稱: " & nm.Name & ", 錯誤內容:" & Err.Description & ", 請確認該條件規則名稱與參照範圍是否正確，若仍無法排除問題，請聯繫 AI&T 同仁。"
    End Function
Function ShowOrHideRows(fieldName As String, relatedRange As String)
    'ex: "B2.YES_and_B3.NO_or_B4.YES__SHOW"
    '將條件分割
    Dim conditionsStr As String
    conditionsStr = Split(fieldName, "__")(0)
    '全形轉半形
    conditionsStr = StrConv(conditionsStr, vbNarrow)
    '小寫轉大寫
    conditionsStr = UCase(conditionsStr)

    Dim actionValue As String
    actionValue = Split(fieldName, "__")(1)
    '若包含"."，則再分割一次，取第一個
    If InStr(1, actionValue, ".") > 0 Then
        actionValue = Split(actionValue, ".")(0)
    End If
    '全形轉半形
    actionValue = StrConv(actionValue, vbNarrow)
    '小寫轉大寫
    actionValue = UCase(actionValue)
    
    '獲取sheetName
    Dim sheetName As String
    Dim sheetNameForRange As String
    sheetNameForRange = Split(relatedRange, "!")(0)
    
    '去除 "="
    sheetName = Replace(sheetNameForRange, "=", "")
    '去除 "'"
    sheetName = Replace(sheetName, "'", "")

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

    '若actionValue為SHOWSHEET或HIDESHEET，則targetRange為當前sheet的A1
    If actionValue = "SHOWSHEET" Or actionValue = "HIDESHEET" Then
        Set targetRange = Range("A1:A1")
    Else
        Set targetRange = Range(sheetNameForRange & "!" & startCell & ":" & endCell)
    End If

    '檢查每個條件, 若有一個條件不符合, 則不顯示
    If CheckCondition(conditionsStr) Then
        If actionValue = "SHOW" Then
            '顯示 Row
            targetRange.EntireRow.Hidden = False
        ElseIf actionValue = "HIDE" Then
            '隱藏 Row
            targetRange.EntireRow.Hidden = True
        ElseIf actionValue = "SHOWSHEET" Then
            '顯示 WorkSheet
            Worksheets(sheetName).Visible = True
        ElseIf actionValue = "HIDESHEET" Then
            '隱藏 WorkSheet
            Worksheets(sheetName).Visible = False
        End If
    ElseIf Not CheckCondition(conditionsStr) Then
        If actionValue = "SHOW" Then
            '隱藏 Row
            targetRange.EntireRow.Hidden = True
        ElseIf actionValue = "HIDE" Then
            '顯示 Row
            targetRange.EntireRow.Hidden = False
        ElseIf actionValue = "SHOWSHEET" Then
            '隱藏 WorkSheet
            Worksheets(sheetName).Visible = False
        ElseIf actionValue = "HIDESHEET" Then
            '顯示 WorkSheet
            Worksheets(sheetName).Visible = True
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

    '將字串內出現的 "_and_" 與 "_or_" 替換成 "|",　"..R.." 與 "..L.." 去除
    inputTmpString = Replace(Replace(Replace(Replace(condition, "..R..", ""), "..L..", ""), "_AND_", "|"), "_OR_", "|")

    '宣告條件陣列
    Dim columnInfoArray() As String
    '將字串以 "|" 分割成陣列
    columnInfoArray = Split(inputTmpString, "|")

    '逐筆檢查條件
    Dim columnInfo As Variant
    For Each columnInfo In columnInfoArray
        ResultCondition = Replace(ResultCondition, columnInfo, CheckFieldValue(columnInfo))
    Next columnInfo
    
    '將字串內出現的 "_and_" 與 "_or_" 替換成 "*" 與 "+"
    ResultCondition = Replace(Replace(ResultCondition, "_AND_", "*"), "_OR_", "+")
    '將字串內出現的 "True" 與 "False" 替換成 "1" 與 "0"
    ResultCondition = Replace(Replace(ResultCondition, "True", "1"), "False", "0")
    '將字串內出現的 "..L.." 替換成 "(" 及 "..R.."替換成 ")"
    ResultCondition = Replace(ResultCondition, "..L..", "(")
    ResultCondition = Replace(ResultCondition, "..R..", ")")
    
    '計算 ResultCondition 的結果
    If Evaluate(ResultCondition) > 0 Then
        ResultBool = True
    Else
        ResultBool = False
    End If
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

    '檢查目標欄位的值是否等於條件欄位的值，或目標欄位包含條件欄位的值
    If targetValue = fieldValue Or InStr(1, targetValue, fieldValue, vbTextCompare) > 0 Then
        CheckFieldValue = "True"
    Else
        CheckFieldValue = "False"
    End If
End Function

'僅在需要大量更改命名區域時使用
Public Sub RescopeNamedRangesToWorksheet()
Dim wb As Workbook
Dim ws As Worksheet
Dim objName As Name
Dim sWsName As String
Dim sWbName As String
Dim sRefersTo As String
Dim sObjName As String
Set wb = ActiveWorkbook
Set ws = ActiveSheet
sWsName = ws.Name
sWbName = wb.Name

'Loop through names in worksheet.
For Each objName In wb.Names
'Check name is visble.
    If objName.Visible = True Then
'Check name refers to a range on the active sheet.
        If InStr(1, objName.RefersTo, sWsName, vbTextCompare) Then
            sRefersTo = objName.RefersTo
            sObjName = objName.Name
'Check name is scoped to the workbook.
            If objName.Parent.Name = sWbName Then
'Delete the current name scoped to workbook replacing with worksheet scoped name.
                objName.Delete
                ws.Names.Add Name:=sObjName, RefersTo:=sRefersTo
            End If
        End If
    End If
Next objName
End Sub
