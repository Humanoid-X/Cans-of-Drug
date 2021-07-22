Set shell = WScript.CreateObject("WSCript.shell")

Dim count
count = 0

Dim okonly, okcancel, abortretryignore, yesnocancel, yesno, retrycancel
Dim critical, question, exclamation, information
Dim ok, cancel, abort, retry, ignore, yes, no

okonly = 0
okcancel = 1
abortretryignore = 2
yesnocancel = 3
yesno = 4
retrycancel = 5

critical = 16
question = 32
exclamation = 48
information = 64

ok = 1
cancel = 2
abort = 3
retry = 4
ignore = 5
yes = 6
no = 7

Heat()

Function Heat()
	Dim rand
	Randomize
	rand = "~" & CStr(Rnd * 5 + 95) & "%"
	
	Dim Text1, Title1, Button, Icon
	Text1 = "Overheat Alert!" & vbCrLf & "You played too much game, Your computer is about to overheat!" & vbCrLf & "Current Heat: " & rand & vbCrLf & "Reboot Computer?"
	Title1 = "Overheat Alert!"
	Button = yesnocancel
	Icon = question
	
	X = MsgBox(Text1, Button + Icon, Title1)
	If(X = cancel Or X = no)Then
		Heat()
		count = count + 1
	Else
		
	End If
End Function