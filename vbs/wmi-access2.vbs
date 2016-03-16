Dim objWMIService

Sub ListHardware()
	Wscript.Echo
	Wscript.Echo "-=[Краткая сводка]================"
	Wscript.Echo

	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
        For Each objItem in colItems
		Wscript.Echo "NumberOfProcessors: " & objItem.NumberOfProcessors
		Wscript.Echo "TotalPhysicalMemory: " & objItem.TotalPhysicalMemory/1024/1024/1024 & " GB"
		Wscript.Echo "Name: " & objItem.Name

	Next

	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
		For Each objItem in colItems
		Wscript.Echo "Name OS: " & objItem.Name
	Next
	

	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_LogicalDisk WHERE DriveType = 3")
		For Each objItem in colItems
		Wscript.Echo "Name: " & objItem.Name
		Wscript.Echo "Size: " & objItem.Size/1024/1024/1024 & " GB"
		Wscript.Echo "FreeSpace: " & objItem.FreeSpace/1024/1024/1024 & " GB"
	Next
	Wscript.Echo
End Sub

Sub ListHardware2()
	Wscript.Echo
	Wscript.Echo "-=[Оборудование]================"
	Wscript.Echo
	Set colItems = objWMIService.ExecQuery("Select * from Win32_PnPEntity")
	For Each objItem in colItems
        	Wscript.Echo "Name: "& objItem.Name
	next
	
End Sub

Sub ListSharedFolders()
	Wscript.Echo
	Wscript.Echo "-=[Сетевые папки]================"
	Wscript.Echo
	For Each colItems in objWMIService.ExecQuery("Select * from Win32_Share")
	Wscript.Echo "Path: " & colItems.Path
	next
End Sub

Sub ListSoftware()
	Wscript.Echo
	Wscript.Echo "-=[Программное обеспечение]================"
	Wscript.Echo
	Set colItems = objWMIService.ExecQuery("Select Name, Caption, Vendor, Version from Win32_Product")
	For Each objItem in colItems
		Wscript.Echo "Name: " & objItem.Name
	Next
End Sub

'-------------------------------------------------------------------------------

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	
' Краткая сводка	
ListHardware

' Оборудование
ListHardware2

'Сетевые папки
ListSharedFolders

'Программное обеспечение
ListSoftware