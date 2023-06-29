# Excel VBA 巨集 - 問答式表單程式
用於在 Excel 中建立具有條件要求的問答式表單。該巨集能根據使用者的回答自動顯示或隱藏特定的問題，讓表單填寫更加靈活和便捷。
![image](https://github.com/JackyLeeDesign/excel-column-display-control-macro-cus/blob/main/DEMO.gif)

## 如何使用
1.	開啟 Excel 檔案，並按下 Alt + F11 開啟巨集編輯器。
2.	在巨集編輯器中，選擇「插入」>「模組」，並將 Form.vb 的程式碼貼上。
3.	回到 Excel 檔案，選擇「開發人員」>「巨集」，並選擇「Form」，按下「執行」。
4.	按下「開始」，並依照提示進行設定。
5.	完成後，按下「確定」，並將表單儲存為 xlsm 檔案。

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
根據動作執行的不同，表示顯示、隱藏指令列或顯示、隱藏指定工作頁。
以下是一個示例，展示如何使用 Excel VBA巨集 - 問答式表單建立來建立問答式表單，可參考 Example 資料夾內之 Form.xlsm 檔案：
共有三個問題：Question1、Question2 和 Question3
其回答欄位分別為：C2、C4 和 C6

* 範例情境：
1.	當 Question1 的回答為 YES 時，顯示 Question2
2.	當 Question2 的回答為 NO 時，隱藏 Question3
3.	當 Question3 的回答為 Complete 時，隱藏工作表 Sheet2
根據上述情境，我們需要分別設定三個命名規則：
Question1 => 命名規則名稱 C2.YES__SHOW，參照範圍選擇 C4：當 C2 等於 "YES" 時，顯示 Question2 (C4)。
Question2 => 命名規則名稱 C4.NO__HIDE，參照範圍選擇 C6：當 C4 等於 "NO" 時，隱藏 Question4 (C6)。
Question3 => 命名規則名稱 C6.Complete__HIDESHEET，參照範圍選擇 Sheet2 之任意欄位：當 C6 等於 "Complete" 時，隱藏工作表 Sheet2。
當使用者輸入回答值時，巨集將自動根據條件設定來顯示或隱藏相關的問題和工作表。
註：Excel詳細操作與設定方式請參考使用說明手冊 (Excel VBA 問答式表單程式使用說明.pptx)。

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
如果您在使用 Excel VBA 巨集 - 問答式表單建立過程中遇到任何問題或需要進一步的支援，可聯繫 AI&T 同仁請求協助。