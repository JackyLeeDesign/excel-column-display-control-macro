' ====================================================================================================
' Author:   Jacky Lee
' Date:     2023/06/01
' Version:  1.0.0
' 本程式碼用於 Excel 表單控制，可依照命名規則，自動控制表單欄位顯示或隱藏。
' This code is used for Excel form control, which can automatically control the display or hide of form fields according to the naming rules.
' ====================================================================================================

' ====================================================================================================
' 將底下 Worksheet_Activate 和  Worksheet_Change 放在需要執行巨集的工作表內。
' Put the following Worksheet_Activate and Worksheet_Change in the worksheet that needs to execute the macro.
' ====================================================================================================
Private Sub Worksheet_Activate()
    CheckAllCellNames
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    CheckAllCellNames
End Sub

' ====================================================================================================
' 將底下程式碼放在需要執行的工作表內，或將其放入模組，這樣可以共用，不用每個工作表都放一次。
' Put the following code in the worksheet that needs to execute the macro, or put it in the module, so that it can be shared without putting it in each worksheet.
' ====================================================================================================
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

        ' 獲取當前工作簿
        ' Get the current workbook
        Set wb = ActiveWorkbook
        ' 獲取當前工作表
        ' Get the current worksheet
        Set ws = wb.ActiveSheet
        ' 獲取當前工作表的所有命名區域
        ' Get all named ranges of the current worksheet
        For Each nm In ws.Names
            If InStr(1, nm.Name, "__") > 0 And InStr(1, nm.Name, ".") > 0 Then
                isDo = True
                ' 若命名區域參照範圍包含多個範圍，則跳出提醒視窗，顯示: nm.Name & " 命名區域之參照範圍包含多個範圍，請重新確認。"
                ' If the named range reference range contains multiple ranges, a reminder window will pop up to show: nm.Name & "The reference range of the named range contains multiple ranges, please check again."
                If InStr(1, nm.RefersTo, "+") > 0 Then
                    MsgBox nm.Name & " 命名區域之參照範圍包含多個範圍，請重新確認。" & vbCrLf & nm.Name & "The reference range of the named range contains multiple ranges, please check again."
                    isDo = False
                End If
                If isDo = True Then
                    ShowOrHideRows nm.Name, nm.RefersTo
                End If
            End If
        Next nm
    Exit Function
    ErrorHandler:
        ' 顯示錯誤訊息，內容可自行修改與調整
        ' Display error message, the content can be modified and adjusted by yourself
        MsgBox "程式在讀取命名規則時發生錯誤，規則名稱: " & nm.Name & ", 錯誤內容:" & Err.Description & ", 請確認該條件規則名稱與參照範圍是否正確，若仍無法排除問題，請聯繫 AI&T 同仁。" & vbCrLf & "An error occurred while the program was reading the naming rules, the rule name: " & nm.Name & ", error content: " & Err.Description & ", please check whether the condition rule name and reference range are correct, if the problem cannot be ruled out, please contact AI&T colleagues."
    End Function

' 依照命名規則，顯示或隱藏欄位
' Show or hide fields according to naming rules
Function ShowOrHideRows(fieldName As String, relatedRange As String)
    'ex: "B2.YES_and_B3.NO_or_B4.YES__SHOW"
    ' 將條件分割
    ' Split the conditions
    Dim conditionsStr As String
    conditionsStr = Split(fieldName, "__")(0)
    ' 全形轉半形
    ' Convert full-width to half-width
    conditionsStr = StrConv(conditionsStr, vbNarrow)
    ' 小寫轉大寫
    ' Convert lowercase to uppercase
    conditionsStr = UCase(conditionsStr)

    Dim actionValue As String
    actionValue = Split(fieldName, "__")(1)
    ' 若包含"."，則再分割一次，取第一個
    ' If it contains ".", then split it again and take the first one
    If InStr(1, actionValue, ".") > 0 Then
        actionValue = Split(actionValue, ".")(0)
    End If
    ' 全形轉半形
    ' Convert full-width to half-width
    actionValue = StrConv(actionValue, vbNarrow)
    ' 小寫轉大寫
    ' Convert lowercase to uppercase
    actionValue = UCase(actionValue)
    
    ' 獲取sheetName
    ' Get sheetName
    Dim sheetName As String
    Dim sheetNameForRange As String
    sheetNameForRange = Split(relatedRange, "!")(0)
    
    ' 去除 "="
    ' Remove "="
    sheetName = Replace(sheetNameForRange, "=", "")
    ' 去除 "'"
    ' Remove "'"
    sheetName = Replace(sheetName, "'", "")

    ' 擷取命名中之起始欄位範圍，存入 startCell，若僅為"=Sheet1!$7:$8"，則為 A7
    ' Extract the starting field range in the naming and store it in startCell. If it is only "=Sheet1!$7:$8", it is A7
    Dim startCell As String
    If InStr(1, Split(relatedRange, "!")(1), ":") > 0 Then
        startCell = Split(Split(relatedRange, "!")(1), ":")(0)
    Else
        startCell = Split(relatedRange, "!")(1)
    End If

    ' 擷取命名中之結束欄位範圍，存入 endCell，若僅為"=Sheet1!$7:$8"，則為 A7，若無則為 startCell
    ' Extract the ending field range in the naming and store it in endCell. If it is only "=Sheet1!$7:$8", it is A7, if not, it is startCell
    Dim endCell As String
    If InStr(1, Split(relatedRange, "!")(1), ":") > 0 Then
        endCell = Split(Split(relatedRange, "!")(1), ":")(1)
    Else
        endCell = startCell
    End If
    
    Dim targetRange As Range

    ' 若actionValue為SHOWSHEET或HIDESHEET，則targetRange為當前sheet的A1
    ' If actionValue is SHOWSHEET or HIDESHEET, then targetRange is A1 of the current sheet
    If actionValue = "SHOWSHEET" Or actionValue = "HIDESHEET" Then
        Set targetRange = Range("A1:A1")
    Else
        Set targetRange = Range(sheetNameForRange & "!" & startCell & ":" & endCell)
    End If

    ' 檢查每個條件, 若有一個條件不符合, 則不顯示
    ' Check each condition, if one condition does not meet, it will not be displayed
    If CheckCondition(conditionsStr) Then
        If actionValue = "SHOW" Then
            ' 顯示 Row
            ' Display Row
            targetRange.EntireRow.Hidden = False
        ElseIf actionValue = "HIDE" Then
            ' 隱藏 Row
            ' Hide Row
            targetRange.EntireRow.Hidden = True
        ElseIf actionValue = "SHOWSHEET" Then
            ' 顯示 WorkSheet
            ' Display WorkSheet
            Worksheets(sheetName).Visible = True
        ElseIf actionValue = "HIDESHEET" Then
            ' 隱藏 WorkSheet
            ' Hide WorkSheet
            Worksheets(sheetName).Visible = False
        End If
    ElseIf Not CheckCondition(conditionsStr) Then
        If actionValue = "SHOW" Then
            ' 隱藏 Row
            ' Hide Row
            targetRange.EntireRow.Hidden = True
        ElseIf actionValue = "HIDE" Then
            ' 顯示 Row
            ' Display Row
            targetRange.EntireRow.Hidden = False
        ElseIf actionValue = "SHOWSHEET" Then
            ' 隱藏 WorkSheet
            ' Hide WorkSheet
            Worksheets(sheetName).Visible = False
        ElseIf actionValue = "HIDESHEET" Then
            ' 顯示 WorkSheet
            ' Display WorkSheet
            Worksheets(sheetName).Visible = True
        End If
    End If
    
End Function
Function CheckCondition(condition As String) As Boolean
    ' 宣告 ResultCondition 為字串
    ' Declare ResultCondition as a string
    Dim ResultCondition As String
    ResultCondition = condition

    ' 宣告 ResultBool 為布林值
    ' Declare ResultBool as a boolean
    Dim ResultBool As Boolean
    
    Dim inputTmpString As String

    ' 將字串內出現的 "_and_" 與 "_or_" 替換成 "|",　"..R.." 與 "..L.." 去除
    ' Replace "_and_" and "_or_" in the string with "|", "..R.." and "..L.." removed
    inputTmpString = Replace(Replace(Replace(Replace(condition, "..R..", ""), "..L..", ""), "_AND_", "|"), "_OR_", "|")

    ' 宣告條件陣列
    ' Declare condition array
    Dim columnInfoArray() As String
    ' 將字串以 "|" 分割成陣列
    columnInfoArray = Split(inputTmpString, "|")

    ' 逐筆檢查條件
    ' Check each condition
    Dim columnInfo As Variant
    For Each columnInfo In columnInfoArray
        ResultCondition = Replace(ResultCondition, columnInfo, CheckFieldValue(columnInfo))
    Next columnInfo
    
    ' 將字串內出現的 "_and_" 與 "_or_" 替換成 "*" 與 "+"
    ' Replace "_and_" and "_or_" in the string with "*" and "+"
    ResultCondition = Replace(Replace(ResultCondition, "_AND_", "*"), "_OR_", "+")
    ' 將字串內出現的 "True" 與 "False" 替換成 "1" 與 "0"
    ' Replace "True" and "False" in the string with "1" and "0"
    ResultCondition = Replace(Replace(ResultCondition, "True", "1"), "False", "0")
    ' 將字串內出現的 "..L.." 替換成 "(" 及 "..R.."替換成 ")"
    ' Replace "..L.." with "(" and "..R.." with ")"
    ResultCondition = Replace(ResultCondition, "..L..", "(")
    ResultCondition = Replace(ResultCondition, "..R..", ")")
    
    ' 計算 ResultCondition 的結果
    ' Calculate the result of ResultCondition
    If Evaluate(ResultCondition) > 0 Then
        ResultBool = True
    Else
        ResultBool = False
    End If
    CheckCondition = ResultBool
End Function
Function CheckFieldValue(columnInfo As Variant) As String
    ' B2.YES
    ' 將字串以 "." 分割成陣列
    ' Split the string into an array with "."
    Dim fieldNameParts() As String
    fieldNameParts = Split(columnInfo, ".")

    Dim columnName As String
    columnName = fieldNameParts(0)
    ' 全形轉半形
    ' Convert full-width to half-width
    columnName = StrConv(columnName, vbNarrow)
    ' 小寫轉大寫
    ' Lowercase to uppercase
    columnName = UCase(columnName)

    Dim fieldValue As String
    fieldValue = fieldNameParts(1)
    ' 全形轉半形
    ' Convert full-width to half-width
    fieldValue = StrConv(fieldValue, vbNarrow)
    ' 小寫轉大寫
    ' Lowercase to uppercase

    ' 獲取目標欄位的值
    ' Get the value of the target field
    targetValue = Range(columnName).Value
    ' 全形轉半形
    ' Convert full-width to half-width
    targetValue = StrConv(targetValue, vbNarrow)
    ' 小寫轉大寫
    ' Lowercase to uppercase
    targetValue = UCase(targetValue)

    ' 檢查目標欄位的值是否等於條件欄位的值，或目標欄位包含條件欄位的值
    ' Check if the value of the target field is equal to the value of the condition field, or if the target field contains the value of the condition field
    If targetValue = fieldValue Or InStr(1, targetValue, fieldValue, vbTextCompare) > 0 Then
        CheckFieldValue = "True"
    Else
        CheckFieldValue = "False"
    End If
End Function


' ====================================================================================================
' 因 Excel 無法大量同時更改命名區域的相關屬性，故使用 VBA 進行大量更改，此部分可視需求自行修改，另外此部分僅在需要大量更改命名區域時使用，且為單獨手動執行，故不需加入到各個工作表中
' Because Excel cannot change the relevant properties of named ranges in bulk at the same time, VBA is used to change them in bulk. This part can be modified as needed. In addition, this part is only used when a large number of named ranges need to be changed, and is executed manually separately, so it does not need to be added to each worksheet
' ====================================================================================================
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

' 以工作表為範圍重新命名命名區域
' Loop through names in worksheet.
For Each objName In wb.Names
    ' 檢查名稱是否為可見
    ' Check name is visble.
    If objName.Visible = True Then
        ' 檢查名稱是否參照到活頁簿上的範圍
        ' Check name refers to a range on the active sheet.
        If InStr(1, objName.RefersTo, sWsName, vbTextCompare) Then
            sRefersTo = objName.RefersTo
            sObjName = objName.Name
            ' 檢查名稱是否為活頁簿範圍
            ' Check name is scoped to the workbook.
            If objName.Parent.Name = sWbName Then
                ' 刪除目前命名區域，並以工作表為範圍新增命名區域
                ' Delete the current name scoped to workbook replacing with worksheet scoped name.
                objName.Delete
                ws.Names.Add Name:=sObjName, RefersTo:=sRefersTo
            End If
        End If
    End If
Next objName
End Sub
