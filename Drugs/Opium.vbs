Set WshShell = WScript.CreateObject("WScript.Shell")

wscript.sleep(600000)

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

For i = 1 To 10000 Step 1
  	wscript.sleep(1000)
	WshShell.SendKeys "{CAPSLOCK}"
Next