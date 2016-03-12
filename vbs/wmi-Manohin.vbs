Set objService = GetObject("winmgmts:\\.\Root\CIMV2")
Dim objWMIService

Sub ListHardware()

	Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor")
      
	For Each objItem in colItems
	Wscript.Echo "processors: " & nmbofproc.Numberofcores
		
	Next
End Sub

