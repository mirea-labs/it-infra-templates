
Dim objWMIService

Sub ListHardware()
	Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")

	For Each objItem in colItems
		Wscript.Echo "NumberOfProcessors: " & objItem.NumberOfProcessors
		Wscript.Echo "TotalPhysicalMemory: " & objItem.TotalPhysicalMemory/1024/1024/1024 & " GB"
		Wscript.Echo "Name: " & objItem.Name
	Next
	
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
		For Each objItem in colItems
		Wscript.Echo "Версия ОС: " & objItem.Name
	Next

	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_LogicalDisk WHERE DriveType = 3")
		For Each objItem in colItems
		Wscript.Echo "Size: " & objItem.Size/1024/1024/1024 & " GB"
		Wscript.Echo "FreeSpace: " & objItem.FreeSpace/1024/1024/1024 & " GB"
	Next
	
	Set colItems = objWMIService.ExecQuery("Select * from Win32_PnPEntity")
		For Each objItem in colItems 
        	Wscript.Echo "Name: "& objItem.Name
    Next
	
End Sub



'-------------------------------------------------------------------------------

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	
' Оборудование	
ListHardware



