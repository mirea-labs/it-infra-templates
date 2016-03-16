Dim objWMIService
strComputer = "."
Set objWMIService = GetObject("winmgmts:"& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_MotherBoardDevice")
	For Each objItem in colItems
	NamePC = objItem.SystemName
	Next
Set SF = CreateObject("Scripting.FileSystemObject")
If SF.FileExists(NamePC & ".txt") Then
	Set File = SF.GetFile(NamePC & ".txt")
	File.Delete
End if
Set File = SF.OpenTextFile(NamePC & ".txt",8,true)

i=0
Sub ListShortInfo()
	Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor")
	For Each objItem in colItems
	i=i+1
	Next
	File.Write("Число процессоров: " & i & VbCrLf) 
	Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory")
	i=0
	For Each objItem in colItems
	i=i+ objItem.Capacity/1024/1024	
	Next
	File.Write("Объем оперативной памяти: " & i & VbCrLf)
	File.Write("Сетевое имя машины:" & NamePC & VbCrLf)
	Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem",,48)
	For Each objItem in colItems
	File.Write("Версия ОС: " & objItem.Caption & " " & objItem.OSArchitecture & VbCrLf) 
	Next
	Set colItems = objWMIService.ExecQuery("Select * from Win32_DiskDrive")
	i=0
	For Each objItem in colItems
		If not objItem.Size = " " Then i=i+ objItem.Size
	Next
	i=i/1024/1024/1024
	File.Write("Полный объем жесткого диска: " & i & " Гб" & VbCrLf) 
	Set colItems = objWMIService.ExecQuery("Select * from Win32_LogicalDisk")
	i=0
	For Each objItem in colItems
		If not objItem.Size = " " Then i=i+ objItem.FreeSpace
	Next
	i=i/1024/1024/1024
	File.Write("Свободный объем жесткого диска: " & i & " Гб" & VbCrLf)  
End Sub

Sub ListHardware()
	Set colItems = objWMIService.ExecQuery("Select * from Win32_PnPEntity")

	For Each objItem in colItems
		File.Write("Описание: " & objItem.Description & VbCrLf)  
		File.Write("Производитель: " & objItem.Manufacturer & VbCrLf) 
		File.Write("Имя: " & objItem.Name & VbCrLf) 
		File.Write("Статус: " & objItem.Status & VbCrLf)  
	Next
End Sub

Sub ListSharedFolders
	Set colItems = objWMIService.ExecQuery("Select * from Win32_Share WHERE Type = 0")
	File.Write("Список сетевых папок" & VbCrLf)
	For Each objItem in colItems
		File.Write("Имя: " & objItem.Name & VbCrLf)
		File.Write("Путь: " & objItem.Path & VbCrLf)
		File.Write("Тип: " & objItem.Type & VbCrLf)
	Next
	Set colItems = objWMIService.ExecQuery("Select * from Win32_Share WHERE Type = 1")
	File.Write("Список принтеров" & VbCrLf)
	For Each objItem in colItems
		File.Write("Имя: " & objItem.Name & VbCrLf)
		File.Write("Путь: " & objItem.Path & VbCrLf)
		File.Write("Тип: " & objItem.Type & VbCrLf)
	Next
End Sub

Sub ListSoftware
	Set colItems = objWMIService.ExecQuery("Select Name, Caption, Vendor, Version from Win32_Product")
	File.Write("Список ПО" & VbCrLf)
	For Each objItem in colItems
		File.Write("Имя: " & objItem.Name & VbCrLf)
		File.Write("Описание: " & objItem.Caption & VbCrLf) 
		File.Write("Производитель: " & objItem.Vendor & VbCrLf) 
		File.Write("Версия: " & objItem.Version & VbCrLf) 
	Next
End Sub

'-------------------------------------------------------------------------------



ListShortInfo
' Оборудование	
ListHardware

'Сетевые папки
ListSharedFolders

'Установленные приложения
ListSoftware
File.Close
MsgBox("Готово!!!")
