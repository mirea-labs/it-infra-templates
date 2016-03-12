Set objService = GetObject("winmgmts:{impersonationLevel=impersonate=impersonate}!\\.\Root\CIMV2")
Dim objWMIService

Sub ListHardware()

	Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor")
      
	'For Each objItem in colItems
	Wscript.Echo "processors: " & nmbofproc.Numberofprocessors
		
	'Next
End Sub

