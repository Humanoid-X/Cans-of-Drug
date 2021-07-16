Set WshShell = WScript.CreateObject("WScript.Shell")
Dim sapi
Set sapi = CreateObject("sapi.spvoice")

Dim rand

For i = 1 To 10 Step 1
	WshShell.run("notepad")
	
	wscript.sleep(200)

	Randomize
	rand = Rnd * 100
	wscript.sleep(rand)
	WshShell.AppActivate("notepad")
	WshShell.SendKeys "A"

	Randomize
	rand = Rnd * 100
	wscript.sleep(rand)
	WshShell.AppActivate("notepad")
	WshShell.SendKeys "b"

	Randomize
	rand = Rnd * 100
	wscript.sleep(rand)
	WshShell.AppActivate("notepad")
	WshShell.SendKeys "i"

	Randomize
	rand = Rnd * 100
	wscript.sleep(rand)
	WshShell.AppActivate("notepad")
	WshShell.SendKeys "h"

	Randomize
	rand = Rnd * 100
	wscript.sleep(rand)
	WshShell.AppActivate("notepad")
	WshShell.SendKeys "e"

	Randomize
	rand = Rnd * 100
	wscript.sleep(rand)
	WshShell.AppActivate("notepad")
	WshShell.SendKeys "n"

	Randomize
	rand = Rnd * 100
	wscript.sleep(rand)
	WshShell.AppActivate("notepad")
	WshShell.SendKeys "s"

	Randomize
	rand = Rnd * 100
	wscript.sleep(rand)
	WshShell.AppActivate("notepad")
	WshShell.SendKeys "a"
Next

sapi.speak("toidi na era uoy lets goooooooooooooooooooooo duxk qell at ledt tye witcer ink can yo roil a dias wayne water steep flooding foam you faucet vee are act de edgy speen tmaime")