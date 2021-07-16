Set WshShell = WScript.CreateObject("WScript.Shell")

Dim rand

For i = 1 To 10 Step 1
	WshShell.run("notepad")
	
	wscript.sleep(1000)

	Randomize
	rand = Rnd * 500
	wscript.sleep(rand)
	WshShell.AppActivate("notepad")
	WshShell.SendKeys "A"

	Randomize
	rand = Rnd * 500
	wscript.sleep(rand)
	WshShell.AppActivate("notepad")
	WshShell.SendKeys "b"

	Randomize
	rand = Rnd * 500
	wscript.sleep(rand)
	WshShell.AppActivate("notepad")
	WshShell.SendKeys "i"

	Randomize
	rand = Rnd * 500
	wscript.sleep(rand)
	WshShell.AppActivate("notepad")
	WshShell.SendKeys "h"

	Randomize
	rand = Rnd * 500
	wscript.sleep(rand)
	WshShell.AppActivate("notepad")
	WshShell.SendKeys "e"

	Randomize
	rand = Rnd * 500
	wscript.sleep(rand)
	WshShell.AppActivate("notepad")
	WshShell.SendKeys "n"

	Randomize
	rand = Rnd * 500
	wscript.sleep(rand)
	WshShell.AppActivate("notepad")
	WshShell.SendKeys "s"

	Randomize
	rand = Rnd * 500
	wscript.sleep(rand)
	WshShell.AppActivate("notepad")
	WshShell.SendKeys "a"
Next