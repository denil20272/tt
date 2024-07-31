9+2.9.10.13.18.25.26
9-7.15.16
10+3.7.8.15.16.23.24.31
10-5.18.26

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



