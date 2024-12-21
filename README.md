




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


