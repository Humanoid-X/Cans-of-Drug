Dim which
Randomize(Rnd()*100)
which = Int(Int((4 - 1 + 1) * Rnd + 1))

If(which = 1)Then
	Hack1()
ELseIf(which = 2)Then
	Hack2()
ElseIf(which = 3)Then
	Hack3()
Else
	Hack4()
End If

Function Hack1()
	Set WshShell = WScript.CreateObject("WScript.Shell")
	Dim sapi
	Set sapi = CreateObject("sapi.spvoice")

	Dim rand

	For i = 1 To 10 Step 1
		WshShell.run("notepad")
		
		wscript.sleep(200)

		Randomize(Rnd()*100)
		rand = Rnd * 500 + 500
		wscript.sleep(rand)
		WshShell.AppActivate("notepad")
		WshShell.SendKeys "A"

		Randomize(Rnd()*100)
		rand = Rnd * 200
		wscript.sleep(rand)
		WshShell.AppActivate("notepad")
		WshShell.SendKeys "b"

		Randomize(Rnd()*100)
		rand = Rnd * 200
		wscript.sleep(rand)
		WshShell.AppActivate("notepad")
		WshShell.SendKeys "i"

		Randomize(Rnd()*100)
		rand = Rnd * 200
		wscript.sleep(rand)
		WshShell.AppActivate("notepad")
		WshShell.SendKeys "h"

		Randomize(Rnd()*100)
		rand = Rnd * 200
		wscript.sleep(rand)
		WshShell.AppActivate("notepad")
		WshShell.SendKeys "e"

		Randomize(Rnd()*100)
		rand = Rnd * 200
		wscript.sleep(rand)
		WshShell.AppActivate("notepad")
		WshShell.SendKeys "n"

		Randomize(Rnd()*100)
		rand = Rnd * 200
		wscript.sleep(rand)
		WshShell.AppActivate("notepad")
		WshShell.SendKeys "s"

		Randomize(Rnd()*100)
		rand = Rnd * 200
		wscript.sleep(rand)
		WshShell.AppActivate("notepad")
		WshShell.SendKeys "a"
	Next

	sapi.speak("toidi na era uoy lets goooooooooooooooooooooo duxk qell at ledt tye witcer ink can yo roil a dias wayne water steep flooding foam you faucet vee are act de edgy speen tmaime")
	sapi.speak(GenerateRandomString(1000))
End Function

Function Hack2()
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
End Function

Function Hack3()
	for i=25 to 0 step -1
		Set oIE = CreateObject("InternetExplorer.Application")
		
		With oIE
			.navigate("about:blank")
			.Document.Title = "Dummy LOL" & string(100, chrb(160))
			.resizable=1
			.height=500
			.width=500
			.menubar=0
			.toolbar=0
			.statusBar=0
			.visible=1
		End With
	next

	Countdown()
End Function

Function Hack4()
	Set WshShell = WScript.CreateObject("WScript.Shell")
	
	wscript.sleep(600000)
	
	For i = 1 To 100 Step 1
		Randomize(Rnd()*100)
		wscript.sleep(Rnd() * 1000)
		wscript.sleep(1000)
		WshShell.SendKeys "w"
		wscript.sleep(1000)
		WshShell.SendKeys "a"
		wscript.sleep(1000)
		WshShell.SendKeys "w"
		wscript.sleep(1000)
		WshShell.SendKeys "a"
		wscript.sleep(1000)
		WshShell.SendKeys "s"
		wscript.sleep(1000)
		WshShell.SendKeys "d"
		wscript.sleep(1000)
		WshShell.SendKeys "s"
		wscript.sleep(1000)
		WshShell.SendKeys "a"
		wscript.sleep(1000)
		WshShell.SendKeys "w"
		wscript.sleep(1000)
		WshShell.SendKeys "w"
		wscript.sleep(1000)
		WshShell.SendKeys "a"
		wscript.sleep(1000)
		WshShell.SendKeys "w"
	Next
	
	For i = 1 To 10000 Step 1
		wscript.sleep(1000)
		WshShell.SendKeys "{CAPSLOCK}"
	Next
End Function

Function GenerateRandomString(StrLen)
	Dim myStr
	Dim random
	Const MainStr= "0123456789abcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyz"
	For i = 1 to StrLen
		Randomize(Rnd()*100)
		random = Int((StrLen - 1 + 1) * Rnd + 1)
		myStr=myStr & Mid(MainStr,random,1)
	Next
	GenerateRandomString = myStr
End Function

Function Heat()
	Dim rand
	Randomize(Rnd()*100)
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

Function Countdown()
	Set oIE = CreateObject("InternetExplorer.Application")

	With oIE
		.navigate("about:blank")
		.Document.Title = "Restart LOL" & string(100, chrb(160))
		.resizable=1
		.height=300
		.width=300
		.menubar=0
		.toolbar=0
		.statusBar=0
		.visible=1
	End With
	
	Do while oIE.Busy
		wscript.sleep 500
	Loop
	
	oIE.document.body.innerHTML = "<div id=""countdown"" style=""font: 36pt sans-serif;text-align:center;""></div>"
	
	for i=10 to 0 step -1
		oIE.document.all.countdown.innerText= "System Restarting in " & i
		wscript.sleep 1000
	next
	oIE.quit
End Function