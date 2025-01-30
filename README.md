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




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
