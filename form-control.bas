' ====================================================================================================
' Author:   Jacky Lee
' Date:     2023/07/25
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
        Dim wb As Workbook
        Dim ws As Worksheet
        Dim nm As Name

        ' 獲取當前工作簿
        ' Get the current workbook
        Set wb = ActiveWorkbook
        ' 獲取當前工作表
        ' Get the current worksheet
        Set ws = wb.ActiveSheet
        ' 獲取當前工作表的所有命名區域
        ' Get all named ranges of the current worksheet

        ' 宣告條件集合
        ' Declare condition collection
        Dim conditions As Collection

        ' 獲取當前邏輯欄
        ' Get the current logic column
        Dim logicColumn As String
        logicColumn = "G"

        ' 讀取邏輯欄所有資料並依序存入條件集合，即使隱藏之邏輯欄也要獲取，有值才存入
        ' Read all data in the logic column and store it in the condition collection in order, even if the hidden logic column must be obtained, and only store it if it has a value
        Set conditions = New Collection
        Dim condition As Variant
        Dim currentValue As String

        For i = 1 To ws.UsedRange.Rows.Count
            'MsgBox Range(logicColumn & i).Value
            currentValue = Range(logicColumn & i).Value
            If currentValue <> "" Then
                ' 將邏輯欄的換行符號移除
                ' Remove the line break symbol in the logic column
                currentValue = Replace(currentValue, Chr(10), "")
                
                ' 若邏輯欄有分號，則再將邏輯欄的值以分號分割，並依序存入條件集合
                ' If the logic column has a semicolon, then split the value of the logic column with a semicolon and store it in the condition collection in order
                If InStr(1, currentValue, ";") > 0 Then
                    Dim tmpStr As String
                    tmpStr = currentValue
                    Dim tmpArray() As String
                    tmpArray = Split(tmpStr, ";")
                    For Each tmp In tmpArray
                        conditions.Add tmp
                    Next tmp
                Else
                    conditions.Add currentValue
                End If
                
            End If
        Next i

        ' 依序處理條件集合中的條件
        ' Process the conditions in the condition collection in order
        For Each condition In conditions
            On Error GoTo ConditionErrorHandler
            ' 檢查條件是否為空值，且必須包含 show 或 hide 字串
            ' Check whether the condition is empty and must contain the string show or hide
            If condition <> "" and (InStr(1, condition, "show") > 0 or InStr(1, condition, "hide") > 0) Then
                ' ex. Q1Answer.YES and Q2Answer.Yes or Q3Answer.YES show Q4 Q5 Q6
                ' ex. Q1Answer.YES and Q2Answer.Yes or Q3Answer.YES hide Q4 Q5 Q6
                ' 將條件全部轉為小寫，全形轉半形
                ' Convert all conditions to lowercase, full-width to half-width
                condition = StrConv(condition, vbNarrow)
                condition = LCase(condition)
                ' 宣告 answerStr 部分與 actionStr 部分
                ' Declare answerStr and actionStr part
                Dim answerStr As String
                Dim actionStr As String

                ' 宣告 顯示sheet或隱藏sheet旗標
                ' Declare show or hide sheet flag
                Dim isShowSheet As Boolean
                Dim isHideSheet As Boolean
                isShowSheet = False
                isHideSheet = False
                ' 宣告 顯示或隱藏動作旗標
                ' Declare show or hide action flag
                Dim isShow As Boolean
                Dim isHide As Boolean
                isShow = False
                isHide = False

                ' 若條件包含 "showsheet"
                ' If the condition contains "showsheet"
                If InStr(1, condition, "showsheet") > 0 Then
                    isShowSheet = True
                ' 若條件包含 "hidesheet"
                ' If the condition contains "hidesheet"
                ElseIf InStr(1, condition, "hidesheet") > 0 Then
                    isHideSheet = True
                ' 若條件包含 "show"
                ' If the condition contains "show"
                ElseIf InStr(1, condition, "show") > 0 Then
                    isShow = True
                ' 若條件包含 "hide"
                ' If the condition contains "hide"
                ElseIf InStr(1, condition, "hide") > 0 Then
                    isHide = True
                End If

                ' 若條件包含 "showsheet",則將條件根據 "showsheet" 分割分別儲存至 answerStr 與 actionStr
                ' If the condition contains "showsheet", then split the condition according to "showsheet" and store it in answerStr and actionStr respectively
                If isShowSheet Then
                    answerStr = Trim(Split(condition, "showsheet")(0))
                    actionStr = Trim(Split(condition, "showsheet")(1))
                ' 若條件包含 "hidesheet"，則條件根據 "hidesheet" 分割分別儲存至 answerStr 與 actionStr
                ' If the condition contains "hidesheet", then split the condition according to "hidesheet" and store it in answerStr and actionStr respectively
                ElseIf isHideSheet Then
                    answerStr = Trim(Split(condition, "hidesheet")(0))
                    actionStr = Trim(Split(condition, "hidesheet")(1))
                ' 若條件包含 "show",則將條件根據 "show" 分割分別儲存至 answerStr 與 actionStr
                ' If the condition contains "show", then split the condition according to "show" and store it in answerStr and actionStr respectively
                ElseIf isShow Then
                    answerStr = Trim(Split(condition, "show")(0))
                    actionStr = Trim(Split(condition, "show")(1))
                ' 若條件包含 "hide"，則條件根據 "hide" 分割分別儲存至 answerStr 與 actionStr
                ' If the condition contains "hide", then split the condition according to "hide" and store it in answerStr and actionStr respectively
                ElseIf isHide Then
                    answerStr = Trim(Split(condition, "hide")(0))
                    actionStr = Trim(Split(condition, "hide")(1))
                End If

                ' 宣告暫存回答字串
                ' Declare temporary answer string
                Dim answerTmpStr As String

                ' 將字串之 'and' 'or' '(' ')' 移除，僅保留回答部分
                ' Remove 'and' 'or' '(' ')' in the string, only keep the answer part
                answerTmpStr = Replace(Replace(Replace(Replace(answerStr, "and", "|"), "or", "|"), "(", ""), ")", "")

                ' 多個空白改為單空白
                ' Multiple spaces to single space
                answerTmpStr = Replace(answerTmpStr, "  ", " ")

                ' 宣告回答陣列
                ' Declare answerArray
                Dim answerArray() As String

                ' 將字串分割成陣列
                ' Split the string into an array
                answerArray = Split(answerTmpStr, "|")

                ' 宣告回答命名規則名稱
                ' Declare answer name rule
                Dim answerNameRule As String

                ' 宣告回答值
                ' Declare answer value
                Dim answerValue As String

                ' 宣告是否達成條件旗標
                ' Declare whether the condition is met
                Dim isDo As Boolean
                isDo = False

                ' 依序處理陣列中的回答
                ' Process the answers in the array in order
                For Each answer In answerArray
                    ' answer 有值才處理
                    ' Process only if answer has value
                    answer = Trim(answer)
                    If answer <> "" Then
                        ' 若 answer 包含 [], 則取出 [] 內字段作為回答值
                        ' If answer contains [], take out the field in []
                        if InStr(1, answer, "[") > 0 Then
                            answerNameRule = Trim(Split(answer, "[")(0))
                            answerValue = Trim(Split(Split(answer, "[")(1), "]")(0))
                        Else
                            ' 根據 ' ' 分割字串，取第一個為命名規則名稱，第二個為回答值
                            ' Split the string according to ' ', take the first one as the naming rule name, and the second one as the answer value
                            answerNameRule = Trim(Split(answer, " ")(0))
                            answerValue = Trim(Split(answer, " ")(1))
                        End If
                       
                        '判斷 answerValue 的值是否與命名規則值相同且 ws.Names(answerNameRule).RefersToRange 是否被隱藏，若值相同且未被隱藏則將 answerStr 內之命名條件取代為 True
                        ' Determine whether the value of answerValue is the same as the naming rule value and whether ws.Names(answerNameRule).RefersToRange is hidden. If the value is the same and not hidden, replace the naming condition in answerStr with True
                        If LCase(StrConv(ws.Names(answerNameRule).RefersToRange.Value,vbNarrow)) = answerValue and ws.Names(answerNameRule).RefersToRange.EntireRow.Hidden = False Then
                            answerStr = Replace(answerStr, answer, "1")
                            ' ex. Q1Answer.YES and ( Q2Answer.Yes or Q3Answer.YES)
                            ' => 1 and (1 or 1)
                        Else
                            answerStr = Replace(answerStr, answer, "0") 
                        End If
                    End If
                Next answer

                ' 將 answerStr 內出現的 "and" 與 "or" 替換成 "*" 與 "+"
                ' Replace "and" and "or" in answerStr with "*" and "+"
                answerStr = Replace(Replace(answerStr, "and", "*"), "or", "+")

                ' 將結果 answerStr 進行邏輯運算
                ' Perform logical operations on the result answerStr
                If Evaluate(answerStr) > 0 Then
                    isDo = True
                    'MsgBox  answerStr & " => " & isDo
                Else
                    isDo = False
                    'MsgBox  answerStr & " => " & isDo
                End if

                ' 宣告暫存動作字串
                ' Declare temporary action string
                Dim actionTmpStr As String

                ' 宣告動作陣列
                ' Declare action array
                Dim actionArray() As String

                ' 將字串以空白分割成陣列
                ' Split the string into an array with spaces
                actionArray = Split(actionStr, ",")

                ' 根據條件是否達成，決定顯示或隱藏
                ' Determine whether to display or hide according to whether the condition is met
                If isDo Then 
                    If isShowSheet Then
                        ' 將該相關的 sheet 顯示
                        ' Display or hide all the reference ranges of the named range
                        For Each action In actionArray
                            action = Trim(action)
                            ' action 有值才處理
                            ' Process only if action has value
                            If action <> "" Then
                                wb.Sheets(action).Visible = True
                            End If
                        Next action
                    ElseIf isHideSheet Then
                        ' 將該相關的 sheet 隱藏
                        ' Display or hide all the reference ranges of the named range
                        For Each action In actionArray
                            action = Trim(action)
                            If action <> "" Then
                                wb.Sheets(action).Visible = False
                            End If
                        Next action
                    ElseIf isShow Then
                        ' 將該命名區域的參照範圍全部顯示
                        ' Display or hide all the reference ranges of the named range
                        For Each action In actionArray
                            action = Trim(action)
                            If action <> "" Then
                                ws.Names(action).RefersToRange.EntireRow.Hidden = False
                            End If
                        Next action
                    ElseIf isHide Then
                        ' 將該命名區域的參照範圍全部隱藏
                        ' Display or hide all the reference ranges of the named range
                        For Each action In actionArray
                            action = Trim(action)
                            If action <> "" Then
                                ws.Names(action).RefersToRange.EntireRow.Hidden = True
                            End If
                        Next action
                    End If
                Else
                    If isShowSheet Then
                        ' 將該相關的 sheet 顯示
                        ' Display or hide all the reference ranges of the named range
                        For Each action In actionArray
                            action = Trim(action)
                            If action <> "" Then
                                wb.Sheets(action).Visible = False
                            End If
                        Next action
                    ElseIf isHideSheet Then
                        ' 將該相關的 sheet 隱藏
                        ' Display or hide all the reference ranges of the named range
                        For Each action In actionArray
                            action = Trim(action)
                            If action <> "" Then
                                wb.Sheets(action).Visible = True
                            End If
                        Next action
                    ElseIf isShow Then
                        ' 將該命名區域的參照範圍全部顯示
                        ' Display or hide all the reference ranges of the named range
                        For Each action In actionArray
                            action = Trim(action)
                            If action <> "" Then
                                ws.Names(action).RefersToRange.EntireRow.Hidden = True
                            End If
                        Next action
                    ElseIf isHide Then
                        ' 將該命名區域的參照範圍全部隱藏
                        ' Display or hide all the reference ranges of the named range
                        For Each action In actionArray
                            action = Trim(action)
                            If action <> "" Then
                                ws.Names(action).RefersToRange.EntireRow.Hidden = False
                            End If
                        Next action
                    End If
                End If
            End If
        Next condition
        Exit Function
    ConditionErrorHandler:
        MsgBox "處理條件 " + condition + " 發生錯誤，錯誤內容:" & Err.Description & ", 請確認該條件規則名稱與參照範圍是否正確，若仍無法排除問題，請聯繫管理人員。" & vbCrLf & "An error occurred while processing the condition '" + condition + "', please check whether the naming rule name and reference range are correct, if the problem cannot be ruled out, please contact the administrator."
        Exit Function
    ErrorHandler:
        MsgBox "判斷表單題目發生錯誤，錯誤內容:" & Err.Description & "，請聯繫管理人員。" & vbCrLf & "An error occurred while processing the form question, please contact the administrator."
        Exit Function
End Function