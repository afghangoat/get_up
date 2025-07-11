'###
'Config
'###
Dim second
second=1000 'MS

Dim minute
minute=60*second

Dim max_sessions
max_sessions=5

for i = 1 To max_sessions
	wscript.sleep 45*minute
	'wshshell.Sendkeys "i"
	
	X=MsgBox("Get your fat ass up! Stretch a bit!",4+64,"Script")
	wscript.sleep 15*minute
	
	Dim objShell: Set objShell = CreateObject("WScript.Shell")
	Select Case objShell.Popup("You can go back now (Yes=stop,No=new session)", 10, "New session?", 1)
	Case -1 
		'Timed Out
	Case 1
		i=max_sessions+1
		'OK Pressed
	Case 2
		'Cancel Pressed
	End Select
Next

Y=MsgBox("Sessions ended! Go for a walk or to the gym!",4+64,"Script")