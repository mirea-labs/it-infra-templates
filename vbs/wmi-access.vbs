
Dim objWMIService

Sub ListHardware()
	Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")

	For Each objItem in colItems
		Wscript.Echo "NumberOfProcessors: " & objItem.NumberOfProcessors
		Wscript.Echo "TotalPhysicalMemory"& objItem.TotalPhysicalMemory
		Wscript.Echo "Name" & objItem.Name
		Wscript.Echo
	Next
End Sub

'-------------------------------------------------------------------------------

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	
' Оборудование	
ListHardware


