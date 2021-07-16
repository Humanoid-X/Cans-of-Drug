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

	' wait for page to load
	Do while oIE.Busy
		wscript.sleep 500
	Loop

	' prepare document body
	oIE.document.body.innerHTML = "<div id=""countdown"" style=""font: 36pt sans-serif;text-align:center;""></div>"

	' display the countdown
	for i=10 to 0 step -1
		oIE.document.all.countdown.innerText= "System Restarting in " & i
		wscript.sleep 1000
	next
	oIE.quit
End Function