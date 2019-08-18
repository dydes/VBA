'从49页开始需要调整，点击的位置不准，第16页不知道为什么漏了

For i= 1 to 51
	'移动到左侧浏览器，激活窗口
	MoveTo 135, 100
	MouseWheel 5
	MoveTo 135, 240
	Delay 500
	
	'按住鼠标选中本页信息
	LeftDown 1
	MoveTo 900, 270
	Delay 1000
	MouseWheel - 5 
	MoveTo 900, 450
	Delay 500
	MoveTo 900, 850
	Delay 500
	LeftUp 1
	Delay 500
	
	'复制
	KeyDown 17, 1
	KeyPress 67, 1
	KeyUp 17, 1
	
	'下一页
	If i <= 3 or i >= 48 
		MoveTo 550, 910
	Else
		MoveTo 560, 910
	End If
	LeftClick 1
	Delay 500
	
	'取消选中
	MoveTo 900, 1000
	LeftClick 1
	
	'移到右侧记事本，激活窗口
	MoveTo 1200, 10
	LeftClick 1
	Delay 500
	
	'回车两次，粘贴
	KeyPress "Enter", 2
	KeyDown 17, 1
	KeyPress 86, 1
	KeyUp 17, 1
	Delay 500
next