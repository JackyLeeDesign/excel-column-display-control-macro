# Excel VBA macro - Q&A form control macro
Excel VBA macro - Q&A form control macro is a handy tool for creating conditional Q&A forms in Excel. The macro can automatically show or hide specific questions based on the user's answers, making the form filling more flexible and convenient.

## How to use
![image](https://github.com/JackyLeeDesign/excel-column-display-control-macro/blob/main/DEMO.gif)

* Example scenario:
1. When the answer to Question1 is YES, show Question2
2. When the answer to Question2 is NO, hide Question4
3. When the answer to Question3 is Complete, hide Sheet2

* Excel settings:
1. First, check whether the designed questions and answers in Excel are correct.
2. Open the Excel \[Name Manager\] dialog box and set naming rules for each question according to the conditional requirements between the questions.
3. Click Add to create a rule for Question1. Set its naming name to C2.Yes__Show.
    According to the above steps, set three naming rules respectively:
    Question1 => Naming rule name C2.YES__SHOW, reference range select C4: When C2 is equal to "YES", show Question2 (C4).
    Question2 => Naming rule name C4.NO__HIDE, reference range select C6: When C4 is equal to "NO", hide Question4 (C6).
    Question3 => Naming rule name C6.Complete__HIDESHEET, reference range select any field of Sheet2: When C6 is equal to "Complete", hide Sheet2.

4. Copy the code in the form-control.bas file to the macro editor of Excel and save it as an Excel workbook with macros enabled.
5. Test whether the problem is displayed or hidden as expected.
    When the user enters the answer value, the macro will automatically show or hide the relevant questions and worksheets according to the condition settings.

## Naming Rule Setting
* Condition Expression:
A complete naming rule contains two parts, "Question Condition" and "Action". For example, a naming rule is A1.Yes__Show, where "A1.Yes" is the "Question Condition" part, and Show is the "Action" part. The two parts are separated by a double underscore. The meaning of A1.Yes__Show is that if the A1 field is Yes, the display action will be executed. As for the target range to be displayed or hidden, it is based on the reference range set by the naming rule in the Name Manager.
Therefore, if another naming rule is "B2.No__Show", and its reference field is B3, it means that if the B2 field is No, the 3rd row to be displayed.
Assuming that another naming rule is C6.YES__Show, and the reference range is D7-D9, it means that when the value of C6 is YES, the 7-9 rows will be displayed.
If you want to set an empty value, you can set the naming rule as D10. NULLVALUE__Show, and the reference range is set to F11-F14, which means that when D10 is empty, rows 11-14 will be displayed.

* Multiple condition settings:
You can also use "and" and "or" to combine multiple conditions, and the priority of "and" is higher than "or".
For example, A1.YES_and_A2.NO_or_A3.YES__show means that when the value of A1 is YES and the value of A2 is NO, or the value of A3 is YES, the reference range will be displayed.
Of course, you can also use ..L.. to represent "(" and ..R.. to represent ")", to deal with more complex conditions.
For example, A1.YES_and_..L..B3.NO_or_B6.YES..R..__show means that when the value of A1 is YES and the value of B3 is NO or the value of B6 is YES, the reference range will be displayed. At this time, the program will give priority to judging the conditions in parentheses.

* Action Execution:
In the name of the naming area, specify the action to be executed when the condition is met. You can use action keywords to represent specific actions. For example, use "SHOW" to display related questions, and use "HIDE" to hide related questions. To display or hide the worksheet, use "SHOWSHEET" or "HIDESHEET".

* Reference Range:
In the reference range of the naming area, select the specified range, which means the execution target of the action.

## Main Function:
1.	Worksheet_Change (ByVal Target As Range)：
Input: Target { Range }
Triggered when editing Excel, read the edited field information (Target), and determine whether the edited range is a single field or multiple fields. If it is a multiple field, an error message will be displayed and returned to the previous step. If it is a single field, call CheckCurrentCellName.

2.	CheckCurrentCellName(Optional ByVal Target As Range = Nothing)：
Input: Target { Range }
Read all the naming rules in the current editing worksheet, call ArrangeNames to classify and sort the naming rules, and pass the sorted naming rules (collection) into SortNm. SortNm will find the problem associated with the user's current editing field and its sub-problem from these naming rules in order and return (collection). Then pass the naming rule collection into ShowOrHideRows to expand and shrink the Excel field.

3.	ArrangeNames (ByVal InputNames As Names) As Collection：
Input: InputNames { Names }
Output: { Collection }
Divide the incoming naming rules into three categories and sort them, and return a sorted Collection. The classification and order are as follows:
1.	The remaining naming rules
2.	The naming rules containing "or" or "and"
3.	The naming rules containing "sheet"

4.	SortNm (ByVal InputNames As Collection) As Collection：
Input: InputNames { Collection }
Output: { Collection }
Sort the naming rules in the collection according to the priority of the naming rules, and return a sorted Collection.

5.	ShowOrHideRows (ByVal InputNames As Collection)：
Input: InputNames { Collection }
Expand or shrink the Excel field according to the naming rules in the collection.

* Other Functions:
1.	GetNmRows(name As String)：
Input: name { String }
Get the row number associated with the rule according to the incoming naming rule.
For example:
Input - "D21.v_or_D32.v_or_D35.v_orD38.v_or_D56.v__show"
Output - Return [21,32,35,38,56], the type is Integer() array

2.  CheckCondition (condition As String) As Boolean
Input: condition { String }
Output: { Boolean }
Check if the condition is met. Pass in the condition of the first half of the naming rule,
For example, the complete naming rule is "A1.Yes_or_B4.v__show", then the condition is "A1.Yes_or_B4.v", pass this section into CheckCondition ("A1.Yes_or_B4.v"), it will The condition is split into A1, Yes and B4.v, and then passed into CheckFieldValue to determine whether the A1 field is "Yes" or the B4 field is "v". The result is replaced back to the original condition. If A1.Yes is True and B4.v is False, the condition is changed from A1.Yes_or_B4.v to =>True or False expression, and finally return the Boolean value of the operation result.

3.  CheckFieldValue(columnInfo As Variant) As String
Input: columnInfo { Variant }
Output: { String }
Check whether the single field condition meets its value. For example, if "A1.Yes" is passed in, the function will obtain the value of the A1 field and determine whether it is "Yes", and the case is not distinguished. If it is "Yes", the Boolean value of the string "True" will be returned. If not, "False" will be returned.

4.  RemoveDuplicate(ByVal coll As Collection) As Collection
Input: coll { Collection }
Output: { Collection }
Remove duplicate naming rules from the naming rule collection, and remove them from the last added item, and return the collection without duplicates.

5. IsInArray(valToBeFound As Variant, arr() As Integer) As Boolean
Input: valToBeFound { Variant }, arr() { Integer }
Output: { Boolean }
Check if a number exists in the array and return a Boolean value.

6. IsStringInCollection(searchString As String, coll As Collection) As Boolean
Input: searchString { String }, coll { Collection }
Output: { Boolean }
Check if a string exists in the collection and return a Boolean value.

7. CreateDictionary(name As String, refersTo As String) As Object
Input: name { String }, refersTo { String }
Output: { Object }
Create a dictionary format from the input name and reference range string and return it.

## Contact Information:
If you have any questions or need further support and other suggestions in the process of using Excel VBA macro - Q&A form creation, or want to write this program with me, please contact me.

# Excel VBA 巨集 - 問答式表單程式
Excel VBA巨集 - 問答式表單建立是一個方便的工具，用於在Excel中建立具有條件要求的問答式表單。該巨集能根據使用者的回答自動顯示或隱藏特定的問題，讓表單填寫更加靈活和便捷。

## 如何使用
![image](https://github.com/JackyLeeDesign/excel-column-display-control-macro/blob/main/DEMO.gif)

* 範例情境：
1.	當 Question1 的回答為 YES 時，顯示 Question2
2.	當 Question2 的回答為 NO 時，隱藏 Question3
3.	當 Question3 的回答為 Complete 時，隱藏工作表 Sheet2

* Excel 設定方式：
1. 首先於 Excel 中確認所設計之問題與回答是否正確。
2. 開啟 Excel \[名稱管理員\] 對話方塊，根據問題之間的條件要求，為每個問題設定命名規則。
3. 點擊新增，建立 Question1 之規則。將其命名名稱設定為 C2.Yes__Show。
    依照上述步驟，分別設定三個命名規則：
    Question1 => 命名規則名稱 C2.YES__SHOW，參照範圍選擇 C4：當 C2 等於 "YES" 時，顯示 Question2 (C4)。
    Question2 => 命名規則名稱 C4.NO__HIDE，參照範圍選擇 C6：當 C4 等於 "NO" 時，隱藏 Question4 (C6)。
    Question3 => 命名規則名稱 C6.Complete__HIDESHEET，參照範圍選擇 Sheet2 之任意欄位：當 C6 等於 "Complete" 時，隱藏工作表 Sheet2。

4. 將 form-control.bas 檔案內之程式碼複製到 Excel 的巨集編輯器中，並儲存成啟用巨集的 Excel 活頁簿。
5. 測試問題是否如期顯示或隱藏。
    當使用者輸入回答值時，巨集將自動根據條件設定來顯示或隱藏相關的問題和工作表。

## 命名規則設定說明
* 條件表達式：
一個完整的命名規則包含「問題條件」與「動作」兩個部分。例如有一個命名規則為 A1.Yes__Show ，其中”A1.Yes” 即為「問題條件」部分，Show 為「動作」部分，兩部分中間會以雙底線來區隔，A1.Yes__Show的意思既為 A1欄位若為Yes，則執行顯示動作，至於顯示或隱藏的目標範圍，即是根據該命名規則於命名管理員內設定的參照範圍來做顯示或隱藏。
故假若另一命名規則為 “B2.No__Show”，其參照欄位為 B3，意即 B2 欄位若為 No，則顯示的第 3 列。
再假設一命名規則為 C6.YES__Show，參照範圍為 D7-D9 表示當 C6 的值等於 YES 時，顯示 7-9 列。
若想設定空值，可設定命名規則為 D10. NULLVALUE__Show，參照範圍設定為F11-F14意即當D10為空值時，將顯示11-14列。

* 多條件設定：
您也可以使用「and」和「or」來組合多個條件，and 的優先權高於 or。
例如， A1.YES_and_A2.NO_or_A3.YES__show 表示當 A1 的值等於 YES 並且 A2 的值等於 NO，或者 A3的值等於 YES 時，顯示參照範圍。
當然也可以使用 ..L.. 表示 「(」與 ..R.. 表示 「) 」，來處理更複雜的條件。
例如， A1.YES_and_..L..B3.NO_or_B6.YES..R..__show 表示當 A1 的值等於 YES 且 B3 的值等於 NO 或 B6 的值等於 YES 時，顯示參照範圍，此時程式將優先判斷括號內之條件。

* 動作執行：
在命名區域的名稱中，指定當條件滿足時要執行的動作。您可以使用動作關鍵字來表示特定的動作。例如，使用「SHOW」表示顯示相關問題，使用「HIDE」表示隱藏相關問題。若要顯示或隱藏工作頁，則使用「SHOWSHEET」或「HIDESHEET」。

* 參照範圍：
在命名區域的參照範圍中，選擇指定範圍，即表示動作的執行目標。

## 邏輯說明
* 主要函式：
1.	Worksheet_Change (ByVal Target As Range)：
輸入: Target { Range }
編輯Excel時觸發，讀取編輯之欄位資訊(Target)，判斷該編輯範圍視為單欄位還是多欄位，若為多欄位則顯示錯誤提示，並返回上一步，若為單欄則呼叫CheckCurrentCellName。

2.	CheckCurrentCellName(Optional ByVal Target As Range = Nothing)：
輸入: Target { Range }
讀取當前編輯工作頁內之所有命名規則，呼叫ArrangeNames將命名規則分類並排序，將整理後之命名規則(集合)傳入SortNm，SortNm將從這些命名規則中依序找出與User當前編輯欄位關聯之問題與其子問題之命名規則並回傳(集合)，接著將命名規則集合傳入ShowOrHideRows進行Excel欄位地展開與縮小。

3.	ArrangeNames (ByVal InputNames As Names) As Collection：
輸入: InputNames { Names }
輸出: { Collection }
將傳入之命名規則分為三類並排序，回傳一個排序後之Collection，分類和順序如下：
1.	剩餘之命名規則
2.	名稱包含 "or" or "and" 之命名規則
3.	名稱包含 "sheet" 之命名規則

4.	SortNm(ByVal AllNames As Collection, ByVal Target As Range) As Collection：
輸入: AllNames { Collection }, Target { Range }
輸出: { Collection }
找出與當前編輯欄位所關聯的問題與其子問題之命名規則，並回傳集合。

5.	ShowOrHideRows(ieldName As String, relatedRange As String)：
輸入: ieldName { String }, relatedRange { String }
依照命名規則，顯示或隱藏問題列。

* 其他函式：
1.	GetNmRows(name As String)：
輸入: name { String }
根據傳入之命名規則，取得該規則關聯之列編號。
例如：
傳入 - "D21.v_or_D32.v_or_D35.v_orD38.v_or_D56.v__show" 
輸出 - 回傳 [21,32,35,38,56]，其型別為Integer() 陣列

2.	CheckCondition (condition As String) As Boolean
輸入: condition { String }
輸出: { Boolean }
檢查條件是否符合，傳入命名規則內前半部分的條件式，
例如完整之命名規則為 “A1.Yes_or_B4.v__show”，則條件式為”A1.Yes_or_B4.v”，將這段傳入 CheckCondition(”A1.Yes_or_B4.v”)，其會將條件式分別拆開為A1,Yes與B4.v，接著逐一傳入CheckFieldValue判斷A1欄位是否為”Yes”或B4欄位是否為”v”，將其結果取代回原條件式，倘若A1.Yes為True，B4.v為False，則條件是將從 A1.Yes_or_B4.v 變更為=>True or False 運算式，最後回傳運算結果之布林值。

3.	CheckFieldValue(columnInfo As Variant) As String
輸入: columnInfo { Variant }
輸出: { String }
檢查單個欄位條件是否符合其值，例如傳入”A1.Yes”，則該函式將獲取A1欄位值，判斷是否為”Yes”，且大小寫不分，若為”Yes”則回傳字串之布林值”True”，若為否則回傳”False”。

4.	RemoveDuplicate(ByVal coll As Collection) As Collection
輸入: coll { Collection }
輸出: { Collection }
移除命名規則集合中重複之命名規則，且從最後新增的項目開始移除，並回傳去除重複之集合。

5.	IsInArray(valToBeFound As Variant, arr() As Integer) As Boolean
輸入: valToBeFound { Variant }, arr() { Integer }
輸出: { Boolean }
檢查某數字是否存在於陣列中，並回傳布林值。

6.	IsStringInCollection(searchString As String, coll As Collection) As Boolean
輸入: searchString { String }, coll { Collection }
輸出: { Boolean }
檢查某字串是否存在於集合中，並回傳布林值。

7.	CreateDictionary(name As String, refersTo As String) As Object
輸入: name { String }, refersTo { String }
輸出: { Object }
將輸入之名稱與參照範圍字串，製作成字典格式並回傳。

## 聯繫資訊：
如果您在使用 Excel VBA 巨集 - 問答式表單建立過程中遇到任何問題或需要進一步的支援與其他建議，或想與我一起撰寫此小程式，歡迎與我聯繫。
