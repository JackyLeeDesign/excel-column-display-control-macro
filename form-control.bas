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
    If Target.Columns.Count = 1 And Target.Rows.Count = 1 Then
        UrModuleName.CheckCurrentCellName Target
    End If
End Sub

' ====================================================================================================
' 將底下程式碼放在需要執行的工作表內，或將其放入模組，這樣可以共用，不用每個工作表都放一次。
' Put the following code in the worksheet that needs to execute the macro, or put it in the module, so that it can be shared without putting it in each worksheet.
' ====================================================================================================
' 主要邏輯:檢查當前命名規則，自動控制該條件之欄位顯示或隱藏。
' Main logic: Check the current naming rules and automatically control the display or hide of the fields of the condition.
Function CheckCurrentCellName(Optional ByVal Target As Range = Nothing)
    On Error GoTo ErrorHandler
    ' 宣告變數
    ' Declare variables
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim nm As name
    Dim TodoNames As New Collection
    Dim nmObj As Object
    Set wb = ActiveWorkbook
    Set ws = wb.ActiveSheet

    ' 整理命名規則，並回傳待處理之命名規則(問題條件)
    ' Organize the naming rules and return the naming rules to be processed (problem conditions)
    Set TodoNames = SortNm(ws.Names, Target)

    ' 依序將問題欄位顯示或隱藏
    ' Show or hide the problem fields in order
    For Each nmObj In TodoNames
        ShowOrHideRows nmObj("Name"), nmObj("RefersTo")
    Next nmObj
    Exit Function
ErrorHandler:
    ' 顯示錯誤訊息，內容可自行修改與調整
    ' Display error message, the content can be modified and adjusted by yourself
    MsgBox "程式在讀取命名規則時發生錯誤，錯誤內容:" & Err.Description & ", 請確認該條件規則名稱與參照範圍是否正確，若仍無法排除問題，請聯繫 AI&T 同仁。" & vbCrLf & "The program encountered an error while reading the naming rules, the error content is:" & Err.Description & ", please check whether the condition rule name and reference range are correct, if the problem cannot be eliminated, please contact AI&T colleagues."
End Function

' 主要邏輯:整理命名規則，並回傳待處理之命名規則(問題條件)
' Main logic: Organize the naming rules and return the naming rules to be processed (problem conditions)
Function SortNm(ByVal AllNames As Names, ByVal Target As Range) As Collection
    ' 宣告變數
    ' Declare variables
    Dim OutputNames As New Collection
    Dim minRow As Integer
    Dim maxRow As Integer
    Dim selfMinRow As Integer
    Dim selfMaxRow As Integer
    Dim nmRows() As Integer
    Dim nmRow As Variant

    ' 取得該問題要顯示或隱藏的欄位範圍
    ' Get the field range to be displayed or hidden
    minRow = Target.Row
    maxRow = Target.Row + Target.Rows.Count - 1

    ' 紀錄過程中有問題之命名規則名稱，並跳過該規則
    ' Record the naming rule name with problems in the process and skip the rule
    Dim errorNames As New Collection

    For Each nm In AllNames
        ' 若命名規則名稱包含在 errorNames Collection 中，則跳過該規則，
        ' If the naming rule name is included in errorNames, skip the rule
        If IsStringInCollection(nm.name, errorNames) Then
            GoTo NextName
        End If

        nmRows = GetNmRows(nm.name)
        ' 檢查該問題要顯示或隱藏的欄位範圍是否包含問題本身，若包含，表示該問題之邏輯設定錯誤，顯示提醒後並跳過該規則
        selfMinRow = Range(nm.refersTo).Row
        selfMaxRow = Range(nm.refersTo).Row + Range(nm.refersTo).Rows.Count - 1
        For i = selfMinRow To selfMaxRow
            If InStr(1, nm.name, "sheet", vbTextCompare) = 0 And IsInArray(i, nmRows) Then
                MsgBox (nm.name & " 該命名規則名稱出現在參照範圍裏面，可能導致無窮迴圈，已自動忽略該規則。")
                ' 將該規則名稱加入 errorNames
                errorNames.Add nm.name
                GoTo NextName
            End If
        Next i
        ' 從當前編輯之題目開始，往下找出待執行之子問題，根據其先後順序加入待做合集 OutputNames (TodoNames)
        ' Starting from the current edited question, find out the sub-problems to be executed below, and add them to the collection OutputNames (TodoNames) to be done according to their order
        For Each nmRow In nmRows
            If nmRow >= minRow And nmRow <= maxRow Then
                If InStr(1, nm.name, "sheet", vbTextCompare) = 0 Then
                    OutputNames.Add CreateDictionary(nm.name, nm.refersTo)
                    ' 判斷該條件是否含有其他子問題，若有，則將子問題加入 OutputNames
                    ' Add sub-problems to OutputNames
                    Dim nmIndex As Variant
                    For Each nmIndex In SortNm(AllNames, Range(nm.refersTo))
                        OutputNames.Add nmIndex
                    Next nmIndex
                ElseIf InStr(1, nm.name, "sheet", vbTextCompare) > 0 Then
                    OutputNames.Add CreateDictionary(nm.name, nm.refersTo)
                End If
            End If
        Next nmRow
NextName:
    Next nm

    ' 去除集合OutputNames內重複之命名規則，若有重複，從最後添加之合集中移除
    ' Remove duplicate naming rules, if there are duplicates, remove them from the last added collection
    ' return the collection
    Set SortNm = RemoveDuplicate(OutputNames)
End Function

'根據 nmName 取得命名規則判斷時所有行數
' Get all row numbers when naming rules are judged according to nmName
Function GetNmRows(name As String) As Integer()
    ' 宣告 nmRowsTmp 陣列，用於暫存命名規則判斷時的所有行數
    Dim nmRowsTmp() As Integer
    ' 使使用 Regex 取出命名規則中的行數 (ex: "D21.v_or_D32.v_or_D35.v_orD38.v_or_D56.v__show" -> "21,32,35,38,56")，並存入 nmRowsTmp 陣列
    ' Normalize the row numbers in the naming rules (ex: "D21.v_or_D32.v_or_D35.v_orD38.v_or_D56.v__show" -> "21,32,35,38,56"), and store them in the nmRowsTmp array
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    ' 條件為：匹配需以1碼字母開頭以及多個數字結尾，但結果僅回傳數字
    ' The condition is: match the need to start with 1 code letter and end with multiple numbers, but only return numbers
    regex.Pattern = "[0-9]+"
    regex.Global = True
    Dim matches As Object
    Set matches = regex.Execute(name)
    
    ReDim nmRowsTmp(0 To matches.Count - 1) ' 設定臨時陣列的大小
    
    Dim i As Integer
    For i = 0 To matches.Count - 1
        nmRowsTmp(i) = CInt(matches(i))
    Next i
    
    ' 將臨時陣列複製到最終的 nmRows 陣列
    Dim nmRows() As Integer
    ReDim nmRows(0 To UBound(nmRowsTmp))
    For i = 0 To UBound(nmRowsTmp)
        nmRows(i) = nmRowsTmp(i)
    Next i
    
    GetNmRows = nmRows
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

' 檢查條件是否符合
' Check if the condition meets
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

' 檢查欄位值是否符合條件
' Check if the field value meets the conditions
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
    If fieldValue = "NULLVALUE" Then
        fieldValue = ""
    End If

    ' 獲取目標欄位的值
    ' Get the value of the target field
    targetValue = Range(columnName).Value
    ' 全形轉半形
    ' Convert full-width to half-width
    targetValue = StrConv(targetValue, vbNarrow)
    ' 小寫轉大寫
    ' Lowercase to uppercase
    targetValue = UCase(targetValue)
    ' 去除空白
    ' Remove spaces
    targetValue = Replace(targetValue, " ", "")
    

    ' 檢查目標欄位的值是否等於條件欄位的值，或目標欄位包含條件欄位的值，且該欄位並非隱藏狀態
    ' Check if the value of the target field is equal to the value of the condition field, or if the target field contains the value of the condition field
    If targetValue = fieldValue Or InStr(1, targetValue, fieldValue, vbTextCompare) > 0 And fieldValue <> "" And targetValue <> "" Then
        If Range(columnName).EntireRow.Hidden = False Then
            CheckFieldValue = "True"
        Else
            CheckFieldValue = "False"
        End If
    Else
        CheckFieldValue = "False"
    End If
End Function

' 移除重複之命名規則
' Remove duplicate naming rules
Function RemoveDuplicate(ByVal coll As Collection) As Collection
    Dim dict As Object
    Dim item As Variant
    Dim key As Variant
    
    Set dict = CreateObject("Scripting.Dictionary")
    For Each item In coll
        dict(item("Name")) = True
    Next item
    
    Set RemoveDuplicate = New Collection
    For Each item In coll
        If dict(item("Name")) Then
            RemoveDuplicate.Add item
            dict(item("Name")) = False ' 將該名稱的索引標記為已添加
        End If
    Next item
End Function

' 檢查陣列內之數字是否存在於另一個陣列中
' Check if the number in the array exists in another array
Function IsInArray(valToBeFound As Variant, arr() As Integer) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

' 檢查一個字串是否存在於 Collection 中
' Check if a string exists in the Collection
Function IsStringInCollection(searchString As String, coll As Collection) As Boolean
    Dim item As Variant
    
    For Each item In coll
        If item = searchString Then
            IsStringInCollection = True
            Exit Function
        End If
    Next item
    
    IsStringInCollection = False
End Function

' Create a dictionary with "Name" and "RefersTo"
Function CreateDictionary(name As String, refersTo As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Name", name
    dict.Add "RefersTo", refersTo
    Set CreateDictionary = dict
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
