' ====================================================================================================
' Author:   Jacky Lee
' Date:     2023/06/01
' Version:  1.0.0
' 本程式碼用於 Excel 表單控制，可依照命名規則，自動控制表單欄位顯示或隱藏。
' This code is used for Excel form control, which can automatically control the display or hide of form fields according to the naming rules.
' ====================================================================================================

' ====================================================================================================
' 將底下 Worksheet_Open 放置於 "This Workbook" 內, Worksheet_Activate 和  Worksheet_Change 放在需要執行巨集的工作表內。
' Put the following Worksheet_Activate and Worksheet_Change in the worksheet that needs to execute the macro.
' ====================================================================================================
Private Sub Worksheet_Open()
    UrModuleName.CheckCurrentCellName
End Sub

Private Sub Worksheet_Activate()
    UrModuleName.CheckCurrentCellName
End Sub

' 當工作表內容有變更時，自動執行 CheckCurrentCellName
Private Sub Worksheet_Change(ByVal Target As Range)
    ' 若僅是刪除或新增欄位，則不執行 CheckCurrentCellName
    ' If it is only to delete or add fields, CheckCurrentCellName will not be executed
    'If Target.Columns.Count = 0 And Target.Rows.Count = 0 Then
    '    Exit Sub
    'End If
    ' 判斷 Target 是否有多欄位變更，若有，則跳出提醒視窗，顯示: "目前版本不支援多欄位變更，請重新確認。"
    ' Determine whether Target has multiple field changes. If so, a reminder window will pop up to show: "The current version does not support multiple field changes, please check again."
    'If Target.Columns.Count > 1 Then
    '    MsgBox "目前版本不支援多欄位變更，請重新確認。" & vbCrLf & "The current version does not support multiple field changes, please check again."
    '    Exit Sub
    'End If
    ' 判斷 Target 是否有多列變更，若有，則跳出提醒視窗，顯示: "目前版本不支援多列變更，請重新確認。"
    ' Determine whether Target has multiple row changes. If so, a reminder window will pop up to show: "The current version does not support multiple row changes, please check again."
    'If Target.Rows.Count > 1 Then
    '    MsgBox "目前版本不支援多列變更，請重新確認。" & vbCrLf & "The current version does not support multiple row changes, please check again."
    '    Exit Sub
    'End If
    
    ' 若僅有單一欄位變更，則執行 CheckCellName
    ' If there is only one field change, execute CheckCellName
    If Target.Columns.Count = 1 And Target.Rows.Count = 1 Then
        UrModuleName.CheckCurrentCellName Target
    End If
End Sub

' ====================================================================================================
' 將底下程式碼放在需要執行的工作表內，或將其放入模組，這樣可以共用，不用每個工作表都放一次。
' Put the following code in the worksheet that needs to execute the macro, or put it in the module, so that it can be shared without putting it in each worksheet.
' ====================================================================================================
' 檢查當前命名規則，自動控制該條件之欄位顯示或隱藏。
' Check the current naming rules and automatically control the display or hide of the fields of the condition.
Function CheckCurrentCellName(Optional ByVal Target As Range = Nothing)
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
        ' 建立存放待做命名規則之Object Array TodoNames, 其包含兩個屬性: Name, RefersTo
        ' Create an Object TodoNames that stores the naming rules to be done, which contains two properties: Name, RefersTo
        Dim TodoNames() As Object
        ' 一開始將 TodoNames 設為空陣列
        ' Set TodoNames to an empty array at the beginning
        ReDim TodoNames(0 To 0) As Object
        ' 執行 SortNm ws.Names, Target, TodoNames，並回傳至 TodoNames
        ' Execute SortNm ws.Names, Target, TodoNames, and return to TodoNames
        ' 顯示ws.Names數量
        ' Display the number of ws.Names
        MsgBox ws.Names.Count
        TodoNames = SortNm(ws.Names, Target, TodoNames)
        ' 逐一將 TodoNames 內容執行 ShowOrHideRows
        ' Execute naming rules in order
        If UBound(TodoNames) > 0 Then
           For i = 1 To UBound(TodoNames)
                ' 執行 ShowOrHideRows
                ' Execute ShowOrHideRows
                ShowOrHideRows TodoNames(i)("Name"), TodoNames(i)("RefersTo")
            Next i
        End If
        Exit Function
ErrorHandler:
        ' 顯示錯誤訊息，內容可自行修改與調整
        ' Display error message, the content can be modified and adjusted by yourself
        MsgBox "程式在讀取命名規則時發生錯誤，錯誤內容:" & Err.Description & ", 請確認該條件規則名稱與參照範圍是否正確，若仍無法排除問題，請聯繫 AI&T 同仁。" & vbCrLf & "The program encountered an error while reading the naming rules, the error content is:" & Err.Description & ", please check whether the condition rule name and reference range are correct, if the problem cannot be eliminated, please contact AI&T colleagues."
    End Function

' 將 TodoNames 進行排序，傳入相關參數，並回傳排序後的 TodoNames
' Sort TodoNames, pass in related parameters, and return the sorted TodoNames
Function SortNm(ByVal AllNames As Names, ByVal Target As Range, ByRef InputNames() As Object) As Object()
    Dim OutputNames() As Object
    Dim minRow As Integer
    Dim maxRow As Integer
    Dim nmRows() As Integer
    Dim tmpNames() As Object
    Dim tmpNamesLength As Integer
    ' 設定暫時變數，用於存放從 Names 中取出的命名規則
    ' Set temporary variables to store the naming rules taken from Names
    ' 若InputNames不為空陣列，則將OutputNames設為InputNames
    If UBound(InputNames) > 0 Then
        ' 將 InputNames 複製至 OutputNames
        ' Copy InputNames to OutputNames
        OutputNames = InputNames
    Else
    ' 若無 InputNames，則將 OutputNames 設為空陣列
    ' If there is no InputNames, set OutputNames to an empty array
        ReDim OutputNames(0 To 0) As Object
        Set OutputNames(0) = CreateObject("Scripting.Dictionary")
        OutputNames(0).Add "Name", ""
        OutputNames(0).Add "RefersTo", ""
    End If
    
    ' 獲取 Target 範圍之起始列數
    ' Get the start row number of the Target range
    
    minRow = Target.Row
    ' 獲取 Target 範圍之結束列數
    ' Get the end row number of the Target range
    
    maxRow = Target.Row + Target.Rows.Count - 1
    ' 依序讀取 AllNames 內的命名規則，若該規則使用GetNmRows取得的行數有包含在 Target 範圍內，則將該規則加入 OutputNames
    ' Read the naming rules in AllNames in order, if the row number of the rule obtained by GetNmRows is included in the Target range, then add the rule to OutputNames

    For Each nm In AllNames
        '宣告 nmRows 陣列，用於存放命名規則判斷時所有行數
        ' Declare the nmRows array to store all row numbers when naming rules are judged
        ' 獲取命名規則中所有行數
        ' Get all row numbers in the naming rule
        nmRows = GetNmRows(nm.Name)
        For Each nmRow In nmRows
            If nmRow >= minRow And nmRow <= maxRow Then
                If InStr(1, nm.Name, "sheet", vbTextCompare) = 0 Then
                    MsgBox(nm.Name)
                    ReDim Preserve OutputNames(0 To UBound(OutputNames) + 1) As Object
                    Set OutputNames(UBound(OutputNames)) = CreateObject("Scripting.Dictionary")
                    OutputNames(UBound(OutputNames)).Add "Name", nm.Name
                    OutputNames(UBound(OutputNames)).Add "RefersTo", nm.RefersTo
                    ' 繼續將該符合條件之規則遞迴執行 SortNm
                    ' Continue to recursively execute SortNm for the rule that meets the condition
                    ' 為避免堆疊空間不足，先算好 SortNm 回傳之陣列長度，再將 SortNm 回傳之陣列加入 OutputNames
                    ' To avoid insufficient stack space, add the array returned by SortNm to OutputNames one by one
                    tmpNames = SortNm(AllNames, Range(nm.RefersTo), OutputNames)
                    tmpNamesLength = UBound(tmpNames)
                    ReDim Preserve OutputNames(0 To UBound(OutputNames) + tmpNamesLength) As Object
                    Dim nmIndex As Integer
                    For nmIndex = 0 To tmpNamesLength
                        Set OutputNames(UBound(OutputNames) - tmpNamesLength + count) = tmpNames(nmIndex)
                    Next nmIndex
                End If
                IF InStr(1, nm.Name, "sheet", vbTextCompare) > 0 Then
                    ReDim Preserve OutputNames(0 To UBound(OutputNames) + 1) As Object
                    Set OutputNames(UBound(OutputNames)) = CreateObject("Scripting.Dictionary")
                    OutputNames(UBound(OutputNames)).Add "Name", nm.Name
                    OutputNames(UBound(OutputNames)).Add "RefersTo", nm.RefersTo
                End If
            End If
        Next nmRow
    Next nm
    ' 將 OutputNames 內 Name 為空值的元素刪除
    ' Delete the elements in OutputNames whose Name is empty
    Dim i As Integer
    For i = UBound(OutputNames) To 1 Step -1
        If OutputNames(i)("Name") = "" Then
            OutputNames = RemoveElement(OutputNames, i)
        End If
    Next i

    ' 將 OutputNames 去除重複，若重複，從最後新增的元素將其刪除
    ' Remove duplicates from OutputNames, if duplicates, delete them from the last added element
    Dim j As Integer
    For i = UBound(OutputNames) To 0 Step -1
        For j = i - 1 To 0 Step -1
            If OutputNames(i)("Name") = OutputNames(j)("Name") Then
                OutputNames = RemoveElement(OutputNames, i)
                Exit For
            End If
        Next j
    Next i

    ' 回傳 OutputNames
    ' Return OutputNames
    SortNm = OutputNames
End Function

'根據 nmName 取得命名規則判斷時所有行數
' Get all row numbers when naming rules are judged according to nmName
Function GetNmRows(Name As String) As Integer()
    'ex: "B2.YES_and_B3.NO_or_B4.YES__SHOW"
    ' 宣告 nmRows 陣列，用於存放命名規則判斷時所有行數
    ' Declare the nmRows array to store all row numbers when naming rules are judged
    Dim nmRows() As Integer
    ' 使使用Regex取出命名規則中的行數 (ex: "B2.YES_and_B3.NO_or_B4.YES__SHOW" -> "2,3,4")，並存入 nmRows 陣列
    ' Normalize the row numbers in the naming rules (ex: "B2.YES_and_B3.NO_or_B4.YES__SHOW" -> "2,3,4"), and store them in the nmRows array
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    '條件為：匹配需以1碼字母開頭以及多個數字結尾，但結果僅回傳數字
    ' The condition is: match the need to start with 1 code letter and end with multiple numbers, but only return numbers
    regex.Pattern = "[0-9]+\."
    regex.Global = True
    Dim matches As Object
    Set matches = regex.Execute(Name)
    Dim i As Integer
    i = 0
    For Each match In matches
        ReDim Preserve nmRows(i)
        '將match.Value中的"."去除，並轉為數字存入nmRows
        ' Remove "." in match.Value and store it in nmRows as a number
        nmRows(i) = CInt(Replace(match.Value, ".", ""))
        i = i + 1
    Next match
    GetNmRows = nmRows
End Function

Function RemoveElement(InputArray() As Object, Index As Integer) As Object()
    ' 宣告 OutputArray 陣列，用於存放刪除元素後之陣列
    ' Declare the OutputArray array to store the array after deleting the element
    Dim OutputArray() As Object
    ' 將 InputArray 複製至 OutputArray
    ' Copy InputArray to OutputArray
    OutputArray = InputArray
    ' 將 OutputArray 中 Index 位置之元素刪除
    ' Delete the element at the Index position in OutputArray
    Dim i As Integer
    For i = Index To UBound(OutputArray) - 1
        Set OutputArray(i) = OutputArray(i + 1)
    Next i
    ReDim Preserve OutputArray(0 To UBound(OutputArray) - 1) As Object
    ' 回傳 OutputArray
    ' Return OutputArray
    RemoveElement = OutputArray
End Function

' 依照命名規則，顯示或隱藏欄位
' Show or hide fields according to naming rules
Function ShowOrHideRows(fieldName As String, relatedRange As String)
    'ex: "B2.YES_and_B3.NO_or_B4.YES__SHOW"
    ' 將條件分割
    ' Split the conditions
    Dim conditionsStr As String
    conditionsStr = Split(fieldName, "__")(0)
    '若包含 "!" ,再以 "!" 分割，取第二個
    ' If it contains "!", then split it again with "!", and take the second one
    If InStr(1, conditionsStr, "!") > 0 Then
        conditionsStr = Split(conditionsStr, "!")(1)
    End If
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
    fieldValue = UCase(fieldValue)

    ' 獲取目標欄位的值
    ' Get the value of the target field
    targetValue = Range(columnName).Value
    ' 全形轉半形
    ' Convert full-width to half-width
    targetValue = StrConv(targetValue, vbNarrow)
    ' 小寫轉大寫
    ' Lowercase to uppercase
    targetValue = UCase(targetValue)

    ' 檢查目標欄位的值是否等於條件欄位的值，或目標欄位包含條件欄位的值，且該欄位並非隱藏狀態
    ' Check if the value of the target field is equal to the value of the condition field, or if the target field contains the value of the condition field
    If (targetValue = fieldValue Or InStr(1, targetValue, fieldValue, vbTextCompare) > 0) And Range(columnName).EntireColumn.Hidden = False Then
        CheckFieldValue = "True"
    Else
        CheckFieldValue = "False"
    End If
End Function

' 檢查所有命名規則
' Check all naming rules
Function CheckAllCellNames(Optional ByVal Target As Range = Nothing)
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
        Dim isDo As Boolean

        ' 獲取當前工作簿
        ' Get the current workbook
        Set wb = ActiveWorkbook
        ' 獲取當前工作表
        ' Get the current worksheet

        ' 依序執行所有工作頁
        ' Execute all worksheets in order
        For Each ws In wb.Worksheets
            ' 獲取當前工作表的所有命名區域
            ' Get all named ranges of the current worksheet
            For Each nm In ws.Names
                If InStr(1, nm.Name, "__") > 0 And InStr(1, nm.Name, ".") > 0 Then
                    isDo = True
                    ' 若命名區域參照範圍包含多個範圍，則跳出提醒視窗，顯示: nm.Name & " 命名區域之參照範圍包含多個範圍，請檢查是否正確。"
                    ' If the named range reference range contains multiple ranges, a reminder window will pop up, showing: nm.Name & "The reference range of the named range contains multiple ranges, please check if it is correct."
                    If InStr(1, nm.RefersTo, ",") > 0 Then
                        MsgBox (nm.Name & " 命名區域之參照範圍包含多個範圍，請檢查是否正確。")
                         isDo = False
                    End If
                    ' 若命名區域參照範圍包含多個工作表，則跳出提醒視窗，顯示: nm.Name & " 命名區域之參照範圍包含多個工作表，請檢查是否正確。"
                    ' If the reference range of the named range contains multiple worksheets, a reminder window will pop up, showing: nm.Name & "The reference range of the named range contains multiple worksheets, please check if it is correct."
                    If InStr(1, nm.RefersTo, "!") > 0 Then
                        MsgBox (nm.Name & " 命名區域之參照範圍包含多個工作表，請檢查是否正確。")
                         isDo = False
                    End If
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
        Next ws
    Exit Function
    ErrorHandler:
        ' 顯示錯誤訊息，內容可自行修改與調整
        ' Display error message, the content can be modified and adjusted by yourself
        MsgBox "程式在讀取命名規則時發生錯誤，規則名稱: " & nm.Name & ", 錯誤內容:" & Err.Description & ", 請確認該條件規則名稱與參照範圍是否正確，若仍無法排除問題，請聯繫 AI&T 同仁。" & vbCrLf & "An error occurred while the program was reading the naming rules, the rule name: " & nm.Name & ", error content: " & Err.Description & ", please check whether the condition rule name and reference range are correct, if the problem cannot be ruled out, please contact AI&T colleagues."
End Function

' ====================================================================================================
' 僅維護時會使用之 Function，不需要加入到每個工作表，當需要執行時貼至巨集後再手動執行
' Function used only when maintaining, do not need to add to each worksheet, when you need to execute, paste to the macro and then execute manually
' ====================================================================================================
' 將 Excel 中之回答全部清除後，執行該程式，會將所有問題根據設定的條件重新顯示
' After clearing all the answers in Excel, execute the program, and all questions will be displayed again according to the set conditions
Public Function ResetAllQuestions()
    CheckAllCellNames
End Function