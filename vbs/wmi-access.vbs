
Dim objWMIService

Sub ListHardware()
	Set colItems = objWMIService.ExecQuery("Select * from Win32_PnPEntity")

	For Each objItem in colItems
		Wscript.Echo "Description: " & objItem.Description
		Wscript.Echo "Manufacturer: " & objItem.Manufacturer
		Wscript.Echo "Name: " & objItem.Name
		Wscript.Echo "Status: " & objItem.Status
		Wscript.Echo
	Next
End Sub

Sub ListSharedFolders
	Set colItems = objWMIService.ExecQuery("Select * from Win32_Share WHERE Type = 0")

	For Each objItem in colItems
		Wscript.Echo "Name: " & objItem.Name
		Wscript.Echo "Path: " & objItem.Path
		Wscript.Echo "Type: " & objItem.Type
		Wscript.Echo
	Next
End Sub

Sub ListSoftware
	Set colItems = objWMIService.ExecQuery("Select Name, Caption, Vendor, Version from Win32_Product")

	For Each objItem in colItems
		Wscript.Echo "Name: " & objItem.Name
		Wscript.Echo "Caption: " & objItem.Caption
		Wscript.Echo "Vendor: " & objItem.Vendor
		Wscript.Echo "Version: " & objItem.Version
		Wscript.Echo
	Next
End Sub

'-------------------------------------------------------------------------------

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	
' Оборудование	
'ListHardware

'Сетевые папки
ListSharedFolders

'Установленные приложения
'ListSoftware

