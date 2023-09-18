# Excel VBA macro - Q&A form control macro
Excel VBA macro - Q&A form control macro is a handy tool for creating conditional Q&A forms in Excel. The macro can automatically show or hide specific questions based on the user's answers, making the form filling more flexible and convenient.

## How to use
![image](https://github.com/JackyLeeDesign/excel-column-display-control-macro/blob/main/DEMO.gif)

* Example scenario:
1. When the answer of Question1 is YES, show Question2
2. When the answer of Question2 is NO, show Question3, Question4, Question5, Question6
3. When the answer of Question3 is Complete, show Sheet2
4. When the answer of Question4 is YES and the answer of Question5 is A or B, show Question6
5. When the answer of Question6 is No, hide Question7

* Excel settings:
1. First, check whether the designed questions and answers in Excel are correct.
2. Open the Excel \[Name Manager\] dialog box and set a name for each question.
    For example: The range of Question1 is A1 => set its name to "Q.1", the reference range is A1,
        The range of Question2 is A2 => set its name to "Q.2", the reference range is A2, and so on.
3. Then set a name for each answer field.
    For example: The answer field of Question1 is C2 => set its name to "A.1", the reference range is C2.
    The answer field of Question2 is C4 => set its name to "A.2", the reference range is C4, and so on.
4. Pick a column as the rule column for display and hide, default is column "G" (if you need to change, you can change the "logicColumn" variable in the vba code by yourself).
5. Fill in the logic one by one:
    When the answer to Question1 is YES, show Question2 => enter "A.1 Yes show Q.2" in the G2 field.
    When the answer to Question2 is NO, show Question3, Question4, Question5, Question6 => enter "A.2 No show Q.3, Q.4, Q.5" in the G4 field.
    When the answer to Question3 is Complete, show Sheet2 => enter "A.3 Yes Complete showsheet Sheet2" in the G6 field.
    When the answer to Question4 is YES and the answer to Question5 is A or B, show Question6 => enter "A.4 Yes and (A.5 B or A.5 A) show Q.6" in the G8 field.
    When the answer to Question6 is No, hide Question7 => enter "A.6 No hide Q.7" in the G15 field.
    If there are multiple logics in the same question, you can use ";" to separate them.
6. Copy the code in the form-control.bas file to the macro editor of Excel and save it as an Excel workbook with macros enabled.
7. Test whether the questions are displayed or hidden as expected.
    When the user enters the answer value, the macro will automatically show or hide the relevant questions and worksheets according to the conditions.

## Question condition setting description
* Multiple condition settings:
You can also use "and" and "or" to combine multiple conditions, and the priority of "and" is higher than "or".
For example, A.1 Yes and A.2 No or A.3 Yes show Q.4 means that when Question1 answers YES and Question2 answers NO, or Question3 answers YES, show Question4.
Of course, you can also use "(" and ")" to handle more complex judgments.
For example, A.4 Yes and (A.5 B or A.5 A) show Q.6 means that when Question4 answers YES and Question5 answers A or B, show Question 6, at this time the program will give priority to judging the conditions in parentheses.

* Action execution:
Use action keywords in the rules to indicate specific actions. For example, use "SHOW" to indicate the display of related questions, and use "HIDE" to indicate the hiding of related questions. To show or hide the worksheet, use "SHOWSHEET" or "HIDESHEET".

## Function description
1. Worksheet_Activate ():
Trigger when switching worksheets, and call CheckAllCellNames ().
2. Worksheet_Change ():
Trigger when editing Excel, and call CheckAllCellNames ().
3. CheckAllCellNames ():
Read the logic column rules in the current editing worksheet (default is "G" column), and automatically show or hide fields and worksheets according to the rules set by the user.

## Contact information:
If you encounter any problems or need further support and other suggestions during the use of Excel VBA macro - Q&A form control macro, or want to write this program with me, please feel free to contact me.

# Excel VBA 巨集 - 問答式表單程式
Excel VBA巨集 - 問答式表單建立是一個方便的工具，用於在Excel中建立具有條件要求的問答式表單。該巨集能根據使用者的回答自動顯示或隱藏特定的問題，讓表單填寫更加靈活和便捷。

## 如何使用
![image](https://github.com/JackyLeeDesign/excel-column-display-control-macro/blob/main/DEMO.gif)

* 範例情境：
1.	當 Question1 的回答為 YES 時，顯示 Question2
2.	當 Question2 的回答為 NO 時，顯示 Question3, Question4, Question5, Question6
3.	當 Question3 的回答為 Complete 時，顯示工作表 Sheet2
4.	當 Question4 的回答為 YES 且 Question5 的回答為 A 或 B 時，顯示 Question6
5.  當 Question6 的回答為 No 時，隱藏 Question7

* Excel 設定方式：
1. 首先於 Excel 中確認所設計之問題與回答是否正確。
2. 開啟 Excel \[名稱管理員\] 對話方塊，為每個題目設定名稱。
    例如:題目1 的範圍為 A1 => 設定其名稱為 "Q.1"，參照範圍為 A1，
        題目2 的範圍為 A2 => 設定其名稱為 "Q.2"，參照範圍為 A2，以此類推。
3. 接著將每個回答欄位也設定名稱。
    例如:題目1 的回答欄位為 C2 => 設定其名稱為 "A.1"，參照範圍為 C2。
    題目2 的回答欄位為 C4 => 設定其名稱為 "A.2"，參照範圍為 C4，以此類推。
4. 挑選某一欄作為顯示隱藏的規則欄，預設為 "G" 欄 (若需變更，可自行更改 vba 程式碼中的 "logicColumn" 變數)。
5. 逐一將邏輯填上:
    當 Question1 的回答為 YES 時，顯示 Question2 => 在 G2 欄位輸入 "A.1 Yes show Q.2"。
    當 Question2 的回答為 NO 時，顯示 Question3, Question4, Question5, Question6 => 在 G4 欄位輸入 "A.2 No show Q.3, Q.4, Q.5"。
    當 Question3 的回答為 Complete 時，顯示工作表 Sheet2 => 在 G6 欄位輸入 "A.3 Yes Complete showsheet Sheet2"。
    當 Question4 的回答為 YES 且 Question5 的回答為 A 或 B 時，顯示 Question6 => 在 G8 欄位輸入 "A.4 Yes and (A.5 B or A.5 A) show Q.6"。
    當 Question6 的回答為 No 時，隱藏 Question7 => 在 G15 欄位輸入 "A.6 No hide Q.7"。
    若同一題中含有多個邏輯，可使用 ";" 區隔。

6. 將 form-control.bas 檔案內之程式碼複製到 Excel 的巨集編輯器中，並儲存成啟用巨集的 Excel 活頁簿。
7. 測試問題是否如期顯示或隱藏。
    當使用者輸入回答值時，巨集將自動根據條件設定來顯示或隱藏相關的問題和工作表。

## 問題條件設定說明
* 多條件設定：
您也可以使用 「and」和 「or」 來組合多個條件，and 的優先權高於 or。
例如， A.1 Yes and A.2 No or A.3 Yes show Q.4 表示當問題1 回答 YES 且 問題2回答 NO，或問題三回答為 YES 時，顯示問題四。
當然也可以使用 「(」 與 「)」 處理更複雜的判斷。
例如， A.4 Yes and (A.5 B or A.5 A) show Q.6 表示當問題4 回答 YES 且 問題5 回答 A 或 B 則顯示問題 6，此時程式將優先判斷括號內之條件。

* 動作執行：
於規則使用動作關鍵字來表示特定的動作。例如，使用「SHOW」表示顯示相關問題，使用「HIDE」表示隱藏相關問題。若要顯示或隱藏工作頁，則使用「SHOWSHEET」或「HIDESHEET」。

## 函式說明
1.	Worksheet_Activate ()：
切換工作頁時觸發，並呼叫CheckAllCellNames()。

2.	Worksheet_Change ()：
編輯Excel時觸發，並呼叫CheckAllCellNames()。

3.	CheckAllCellNames ()：
讀取當前編輯工作頁內之邏輯欄規則(預設為 ”G” 欄)，根據User設定之規則自動顯示或隱藏欄位及工作頁。 

## 聯繫資訊：
如果您在使用 Excel VBA 巨集 - 問答式表單建立過程中遇到任何問題或需要進一步的支援與其他建議，或想與我一起撰寫此小程式，歡迎與我聯繫。