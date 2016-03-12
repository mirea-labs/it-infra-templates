Dim objWMIService

Sub ListHardware()
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
        For Each objItem in colItems
		Wscript.Echo "NumberOfProcessors: " & objItem.NumberOfProcessors
		Wscript.Echo "TotalPhysicalMemory: " & objItem.TotalPhysicalMemory/1024/1024/1024
		Wscript.Echo "Name: " & objItem.Name

	Next
	
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_DiskDrive")
		For Each objItem in colItems
		Wscript.Echo "Size: " & objItem.Size
	Next
End Sub

'-------------------------------------------------------------------------------

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	
' Оборудование	
ListHardware