# Excel VBA Macro - Questionnaire Form Builder
It is a convenient tool for creating conditional questionnaire forms in Excel. The macro can automatically show or hide specific questions according to the user's answers, making the form filling more flexible and convenient.

![image](https://github.com/JackyLeeDesign/excel-column-display-control-macro/blob/main/DEMO.gif)

## How to use
Use the Name Manager of Excel to create a named range, and the setting description of the named range is as follows:

* Naming rules:
The name of the named range needs to follow a specific naming rule. The name of the named range should contain two underscores (__) and end in the format of "question condition__action". For example, condition 1__Show or condition 2__Hide.

* Condition expression:
The condition part in the name of the named range is used to specify the condition for displaying or hiding the question.
For example, the name of the naming rule is A1.YES__Show, and the reference range is A7-A9, which means that when the value of A1 is equal to YES, rows 7-9 are displayed.

* Condition setting:
According to the condition requirements between questions, specify the conditions for each question of the named range.
You can use "and" and "or" to combine multiple conditions, and the priority of "and" is higher than "or".
For example, A1.YES_and_A2.NO_or_A3.YES__show means that when the value of A1 is equal to YES and the value of A2 is equal to NO, or the value of A3 is equal to YES, the reference range is displayed.

Of course, you can also use ..L.. to represent "(" and ..R.. to represent ")", to deal with more complex conditions.
For example, A1.YES_and_..L..B3.NO_or_B6.YES..R..__show means that when the value of A1 is equal to YES and the value of B3 is equal to NO or the value of B6 is equal to YES, the reference range is displayed. At this time, the program will give priority to judging the conditions in parentheses.

* Action execution:
In the name of the named range, specify the action to be executed when the condition is met. You can use action keywords to represent specific actions. For example, use "SHOW" to display related questions, and use "HIDE" to hide related questions. To display or hide the worksheet, use "SHOWSHEET" or "HIDESHEET".

* Reference range:
In the reference range of the named range, select the specified range, which means the execution target of the action.
Depending on the different actions, it means displaying or hiding the specified row or displaying or hiding the specified worksheet.

## Example and Demo
Here is an example to show how to use Excel VBA Macro - Questionnaire Form Builder to create a questionnaire form. You can refer to the file Form.xlsm in the Example folder.
There are three questions: Question1, Question2 and Question3
The answer fields are: C2, C4 and C6
Example Scenario:
When the answer of Question1 is YES, show Question2
When the answer of Question2 is NO, hide Question4
When the answer of Question3 is Complete, hide Sheet2

According to the above scenario, we need to set three naming rules respectively:
Question1 => Naming rule name C2.YES__SHOW, the reference range is selected as C4: When C2 is equal to "YES", show Question2 (C4).
Question2 => Naming rule name C4.NO__HIDE, the reference range is selected as C8: When C4 is equal to "NO", hide Question2 (C8).
Question3 => Naming rule name C6.Complete__HIDESHEET, the reference range is selected as any field of Sheet2: When C6 is equal to "Complete", hide Sheet2.
When the user enters the answer value, the macro will automatically show or hide the relevant questions and worksheets according to the condition settings.

## Worksheet Trigger Event
Worksheet_Activate: Triggered when the worksheet is activated, used to check all named ranges.
Worksheet_Change: Triggered when the worksheet content changes, also used to check all named ranges.

## Function:
CheckAllCellNames: The main function is used to check all named ranges.
ShowOrHideRows: Show or hide rows or worksheets according to the condition settings.
CheckCondition: Function to check if the condition is met.
CheckFieldValue: Function to check if the value of the target field matches the condition field.
RescopeNamedRangesToWorksheet: Reset the scope of the named range to the worksheet.

## Function Description:
CheckAllCellNames:
Function: Check all named ranges and show or hide the corresponding rows or worksheets according to the condition settings.
Parameter: None.
Return value: None.

ShowOrHideRows:
Function: Show or hide rows or worksheets according to the condition settings.
Parameter:
fieldName: The name of the named range.
relatedRange: The range referenced by the named range.
Return value: None.

CheckCondition:
Function: Check if the condition is met.
Parameter: condition (condition expression).
Return value: Boolean value (True or False).

CheckFieldValue:
Function: Check if the value of the target field matches the condition field.
Parameter: columnInfo (field information).
Return value: Boolean value (True or False).

RescopeNamedRangesToWorksheet:
Function: Reset the scope of the named range to the worksheet.
Parameter: None.
Return value: None.

## How to run the macro manually:
1.	Select the worksheet where you want to run the macro.
2.	In the Excel menu, click the Developer tab.
3.	In the Code group, click Macros.
4.	In the Macro dialog box, select the macro you want to run (for example, CheckAllCellNames).

## Contact Information:
If you have any questions or need further support and other suggestions in the process of using Excel VBA Macro - Questionnaire Form Builder, or want to write this program with me, please feel free to contact me.

# Excel VBA 巨集 - 問答式表單程式
是一個方便的工具，用於在Excel中建立具有條件要求的問答式表單。該巨集能根據使用者的回答自動顯示或隱藏特定的問題，讓表單填寫更加靈活和便捷。

## 如何使用
使用 Excel 之名稱管理員建立命名區域，命名區域的設定說明如下：

* 命名規則：
命名區域的名稱需要遵循特定的命名規則。命名區域的名稱應包含兩個底線（__）並以「問題條件__動作」的格式結束。例如，條件 1__Show 或 條件 2__Hide。
 
* 條件表達式：
命名區域的名稱中的條件部分是用來指定顯示或隱藏該問題的條件。
例如， 命名規則名稱為A1.YES__Show，參照範圍為 A7-A9 表示當 A1 的值等於 YES 時，顯示 7-9 列。

* 條件設定：
根據問題之間的條件要求，為每個問題的命名區域指定條件。
您可以使用「and」和「or」來組合多個條件，and 的優先權高於 or。
例如， A1.YES_and_A2.NO_or_A3.YES__show 表示當 A1 的值等於 YES 並且 A2 的值等於 NO，或者 A3的值等於 YES 時，顯示參照範圍。

當然也可以使用 ..L.. 表示 「(」與 ..R.. 表示 「) 」，來處理更複雜的條件。
例如， A1.YES_and_..L..B3.NO_or_B6.YES..R..__show 表示當 A1 的值等於 YES 且 B3 的值等於 NO 或 B6 的值等於 YES 時，顯示參照範圍，此時程式將優先判斷括號內之條件。

* 動作執行：
在命名區域的名稱中，指定當條件滿足時要執行的動作。您可以使用動作關鍵字來表示特定的動作。例如，使用「SHOW」表示顯示相關問題，使用「HIDE」表示隱藏相關問題。若要顯示或隱藏工作頁，則使用「SHOWSHEET」或「HIDESHEET」。

* 參照範圍：
在命名區域的參照範圍中，選擇指定範圍，即表示動作的執行目標。
根據動作執行的不同，表示顯示、隱藏指令列或顯示、隱藏指定工作頁。

## 示例和示範：
以下是一個示例，展示如何使用 Excel VBA巨集 - 問答式表單建立來建立問答式表單，可參考 Example 資料夾內之 Form.xlsm 檔案：
共有三個問題：Question1、Question2 和 Question3
其回答欄位分別為：C2、C4 和 C6
範例情境：
當 Question1 的回答為 YES 時，顯示 Question2
當 Question2 的回答為 NO 時，隱藏 Question4
當 Question3 的回答為 Complete 時，隱藏工作表 Sheet2

根據上述情境，我們需要分別設定三個命名規則：
Question1 => 命名規則名稱 C2.YES__SHOW，參照範圍選擇 C4：當 C2 等於 "YES" 時，顯示 Question2 (C4)。
Question2 => 命名規則名稱 C4.NO__HIDE，參照範圍選擇 C8：當 C4 等於 "NO" 時，隱藏 Question4 (C8)。
Question3 => 命名規則名稱 C6.Complete__HIDESHEET，參照範圍選擇 Sheet2 之任意欄位：當 C6 等於 "Complete" 時，隱藏工作表 Sheet2。
當使用者輸入回答值時，巨集將自動根據條件設定來顯示或隱藏相關的問題和工作表。

## 工作表觸發事件
Worksheet_Activate：當工作表被啟用時觸發，用於檢查所有的命名區域。
Worksheet_Change：當工作表內容發生變化時觸發，同樣用於檢查所有的命名區域。

## 函數：
CheckAllCellNames：主要函數，用於檢查所有的命名區域。
ShowOrHideRows：根據條件設定來顯示或隱藏行或工作表。
CheckCondition：檢查條件是否符合的函數。
CheckFieldValue：檢查目標欄位的值是否符合條件欄位的函數。
RescopeNamedRangesToWorksheet：重新將命名區域的作用範圍設定為工作表。

## 函數說明：
CheckAllCellNames：
功能：檢查所有的命名區域並根據條件設定顯示或隱藏相應的行或工作表。
參數：無。
返回值：無。

ShowOrHideRows：
功能：根據條件設定來顯示或隱藏行或工作表。
參數：
fieldName：命名區域的名稱。
relatedRange：命名區域參照的範圍。
返回值：無。

CheckCondition：
功能：檢查條件是否符合。
參數：condition（條件表達式）。
返回值：布林值（True或False）。

CheckFieldValue：
功能：檢查目標欄位的值是否符合條件欄位。
參數：columnInfo（欄位訊息）。
返回值：布林值（True或False）。

RescopeNamedRangesToWorksheet：
功能：重新將命名區域的作用範圍設定為工作表。
參數：無。
返回值：無。

## 手動執行巨集的步驟：
1.	選擇需要執行巨集的工作表。
2.	在Excel菜單中的「開發人員」選項卡中，點擊「巨集」按鈕。
3.	在彈出的巨集對話框中，選擇要執行的巨集（例如CheckAllCellNames）。
4.	點擊「執行」按鈕開始執行巨集。
巨集將根據設定的條件和回答值，自動顯示或隱藏相關的問題或工作表。

## 聯繫資訊：
如果您在使用 Excel VBA 巨集 - 問答式表單建立過程中遇到任何問題或需要進一步的支援與其他建議，或想與我一起撰寫此小程式，歡迎與我聯繫。