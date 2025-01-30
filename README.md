Press CTRL
Send C
Release CTRL

Send>^c
WaitClipBoard



10
本26100
輪8000
全2000
績1434
環2500
40小13345
勞1100
健822
福131
餐788
to50538
11
本26100
輪8000
全2000
績1461
環2500
80小26707
國定11小3672
緊急停工10小3338
勞1100
健859
福131
餐1121
to70567
12
本26100
輪1500
全2000
績1429
環2500
33.5小9360
營運組織特別獎金2000
其他獎金100
勞1100
健943
福131
餐1450
to41365
1
本26100
輪1500
全2000
績1385
環2500
40小9859
勞1145
健896
福131
餐1283
to39889





1
15.16.19.20.27.28.31
2
1.5.24.25.28
3
1.4.5.17.24.25.28


Sub DeleteDuplicatesInColumnD()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cell As Range
    Dim uniqueValues As Object
    
    ' 設定要操作的工作表
    Set ws = ThisWorkbook.Sheets("Sheet1") '根據您的實際工作表名稱調整
    
    ' 獲取 D 列最後一行的行號
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    
    ' 使用字典來儲存唯一值
    Set uniqueValues = CreateObject("Scripting.Dictionary")
    
    ' 從最後一行開始遍歷資料，往上檢查重複項
    For i = lastRow To 1 Step -1
        Set cell = ws.Cells(i, "D")
        
        ' 若值未在字典中則新增，若已存在則刪除該列
        If Not uniqueValues.exists(cell.Value) Then
            uniqueValues.Add cell.Value, Nothing
        Else
            cell.EntireRow.Delete
        End If
    Next i
End Sub





Dim lastRow As Long
    Dim currentColumn As Long
    
    currentColumn = ActiveCell.Column ' 取得當前選取的列
    lastRow = Cells(Rows.Count, currentColumn).End(xlUp).Row ' 找到當前列最後一個有內容的行
    
    ' 只在當前列中選取有內容的儲存格
    Range(Cells(1, currentColumn), Cells(lastRow, currentColumn)).SpecialCells(xlCellTypeConstants).Select

Excel快捷鍵
C+A V貼上文本
W+L 鎖定計算機
W+M 縮小所有視窗
W+E 檔案總管
W+D 顯示桌面

按Windows 標誌鍵 + Shift + S 以進行靜態影像剪取
按Print Screen (PrtSc)以進行靜態影像剪取


2024/10/02 19:20 - 2024/10/03 07:20
10.00
1002 山陀兒颱風

2024/10/02 00:00 - 2024/10/02 07:20
5.50
1002 山陀兒颱風

假日50小時夜 16694
緊急停工10小時夜 3339


1.5.8
macro scheduler loop
Shift + Esc
------------
//VK119=F8
OnEvent>Key_Down,VK119,0,STOP

SRT>STOP
  MDL>Stopping
  Exit>0
END>STOP


Label>start
  my code here
  my code here
  my code here
GoTo>start
-----------
Selenium
---------
javascript:(function() { function R(a){ona = "on"+a; if(window.addEventListener) window.addEventListener(a, function (e) { for(var n=e.originalTarget; n; n=n.parentNode) n[ona]=null; }, true); window[ona]=null; document[ona]=null; if(document.body) document.body[ona]=null; } R("contextmenu"); R("click"); R("mousedown"); R("mouseup"); R("selectstart");})()
----------
單月加班上限可至54小時
每三個月加班總時數上限138小時
1個月加班時數上限得放寬為54小時，惟3個月之加班總時數仍不得超過138小時(46小時x3個月)

不計入加班時數
雙週<=80小時之排班



輪班津貼
凡因請事、病假未能在該班次工作，自當月輪班津貼中，每小時扣除1/240的輪班津貼。
工作環境津貼
按實際在該工作站之在職天數發給。事、病假將依實際時數扣除，每小時扣除1/240。
全勤獎金
每月請病假及事假者，缺勤一天(

1/1
除夕
春節農曆1.2.3日
2/28
兒童節
清明節
1/2
3/29

目前公司月投保薪資是採每月依最近三個月的平均薪資調整"月投保薪資"。(例如：4月份投保薪資為1、2、3月的平均薪資)
"平均薪資"之計算基準為前三個月的平均薪資，含每月經常性會發生的工資(但需扣除缺勤扣款)如：底薪、績效獎金、訓練員獎金、全勤獎金、加班費、不休假代金等。
註：投保薪資不含非經常性薪資（如年終獎金、季獎金等）。

輪班津貼
未滿一個月者，每日以1/30比例計發，有請病.事.家庭照顧.曠職，則以實際時數扣除，每小時扣除1/240
全勤獎金
事.病假視為缺勤，缺勤一天內"包含"即扣除1/3，一天以上、二天以內"包含"扣除2/3，二天以上就不發。曠職也不發。



powershell

# 設定圖片檔案路徑
$imagePath = "$env:USERPROFILE\Desktop\image.jpg"

# 讀取圖片檔案並轉換為 Base64
$base64String = [Convert]::ToBase64String((Get-Content -Path $imagePath -Encoding Byte))

# 輸出 Base64 編碼
$base64String > "$env:USERPROFILE\Desktop\image_base64.txt"

# 選擇是否顯示在 PowerShell 上
$base64String



# 設定 Base64 編碼檔案路徑和輸出圖片檔案名稱
$base64FilePath = "$env:USERPROFILE\Desktop\image_base64.txt"
$outputImagePath = "$env:USERPROFILE\Desktop\decoded_image.jpg"

# 讀取 Base64 編碼
$base64String = Get-Content -Path $base64FilePath

# 將 Base64 字串轉回位元組
$imageBytes = [Convert]::FromBase64String($base64String)

# 將位元組寫回為圖片檔案
[System.IO.File]::WriteAllBytes($outputImagePath, $imageBytes)

Write-Host "圖片已成功解碼，儲存為 $outputImagePath"

在 Macro Scheduler 中，可以使用 Repeat 和 Until 來建立迴圈，或使用 While 和 EndWhile。

以下是兩種寫法的範例：

1. 使用 Repeat ... Until 建立迴圈

Let>Counter=0

Repeat
  Let>Counter=Counter+1
  MessageModal>這是第 %Counter% 次執行
Until>Counter=5

說明：
	•	Let>Counter=0：設定計數器初始值為 0。
	•	Repeat：開始迴圈。
	•	Let>Counter=Counter+1：每次執行迴圈時計數器加 1。
	•	MessageModal：彈出訊息框顯示計數。
	•	Until>Counter=5：當計數器等於 5 時結束迴圈。

2. 使用 While ... EndWhile 建立條件迴圈

Let>Counter=0

While>Counter<5
  Let>Counter=Counter+1
  MessageModal>這是第 %Counter% 次執行
EndWhile

說明：
	•	While>Counter<5：當計數器小於 5 時執行迴圈。
	•	EndWhile：結束迴圈。

兩者的執行結果相同，都會執行 5 次迴圈。如果你的邏輯需要更複雜的條件，也可以在 Until 或 While 裡面加入更多判斷條件。

如果有更具體的需求，可以再提供細節，我幫你修改！





Macro Scheduler 6.2 是比較早期的版本，可能不支援新的語法（例如 Repeat ... Until 和 While ... EndWhile）。但你可以使用 Goto 和 Label 來模擬迴圈功能。以下是適用於 Macro Scheduler 6.2 的替代方法：

1. 使用 Goto 和 Label 模擬迴圈

Let>Counter=0
Label>LoopStart
Let>Counter=Counter+1
MessageModal>這是第 %Counter% 次執行
If>Counter<5,Goto>LoopStart

說明：
	•	Let>Counter=0：設定計數器初始值為 0。
	•	Label>LoopStart：定義一個標籤，作為迴圈的起始位置。
	•	If>Counter<5,Goto>LoopStart：如果計數器小於 5，跳回 LoopStart 位置，繼續執行迴圈。

2. 無條件執行固定次數的迴圈

如果你只需要執行固定次數，可以簡化為：

Let>Counter=0
Label>LoopStart
Let>Counter=Counter+1
MessageModal>迴圈執行中，次數：%Counter%
Goto>EndLoop,Counter=5
Goto>LoopStart
Label>EndLoop

說明：
	•	當 Counter=5 時，跳到標籤 EndLoop 結束迴圈。

如果你的版本限制更多，或需要特定功能，可以再說明需求，我幫你設計替代方案！
