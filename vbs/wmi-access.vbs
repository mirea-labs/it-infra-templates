Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")

For Each objMoth In objWMIService.ExecQuery("SELECT * FROM Win32_MotherboardDevice")
	PC_Name = objMoth.SystemName
Next

set FSO = CreateObject("Scripting.FileSystemObject")
If FSO.FileExists(PC_Name & ".txt") Then 
	Set file = FSO.GetFile(PC_Name & ".txt")
	File.Delete
End If
set File = FSO.OpenTextFile(PC_Name & ".txt", 8, True)

i=0
	
'dim text
For Each objProc In objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
	i=i+1
Next
Print "Краткая сводка:" & vbCrlf & vbCrlf
Print "Кол-во процессоров - " & i

i=0
For Each objPhMem In objWMIService.ExecQuery("SELECT * FROM Win32_PhysicalMemory")
	i = i+ objPhMem.Capacity/1024/1024
Next
Print vbCrlf & "Объем оперативной памяти - " & i & " Мегабайт"

Print vbCrlf & "Сетевое имя машины - " & PC_Name

For Each objItem in objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem",,48) 
	Print vbCrlf & "Версия ОС - " & objItem.Caption & " " & objItem.OSArchitecture
Next

i=0
For Each objDisk In objWMIService.ExecQuery("SELECT * FROM Win32_DiskDrive")
	if Not objDisk.Size = "" Then i = i + objDisk.Size
Next
i = i/1024/1024/1024
Print vbCrlf & "Полный объем жестких дисков - " & i & " Гигабайт"

i=0
For Each objDisk In objWMIService.ExecQuery ("Select * From Win32_LogicalDisk")
	if Not objDisk.Size = "" Then i = i + objDisk.FreeSpace
Next
i = i/1024/1024/1024
Print vbCrlf & "Свободный объем жестких дисков - " & i & " Гигабайт" & vbCrlf & "###########################################################################################################" & vbCrlf & "Оборудование:" & vbCrlf

For Each objItem in objWMIService.ExecQuery("Select * from Win32_PnPEntity")
	Print vbCrlf & "Описание: " & objItem.Description
	Print vbCrlf & "Производитель: " & objItem.Manufacturer
	Print vbCrlf & "Имя: " & objItem.Name
	Print vbCrlf & "Статус: " & objItem.Status & vbCrlf
Next
Print "###########################################################################################################" & vbCrlf & "Ресурсы:" & vbCrlf

For Each objItem in objWMIService.ExecQuery("Select * from Win32_Share where type<>1")
	Print vbCrlf & objItem.Caption & " " & objItem.Name
Next
Print vbCrlf
For Each objItem in objWMIService.ExecQuery("Select * from Win32_Share where type=1")
	Print vbCrlf & objItem.Caption & " " & objItem.Name
Next

Print "###########################################################################################################" & vbCrlf & "Программное обеспечение:" & vbCrlf
For Each objItem in objWMIService.ExecQuery("Select * from Win32_Product")
	Print vbCrlf & objItem.Name & objItem.Version
Next

Sub Print(text)
	'WScript.Echo text
	File.Write(text)
End Sub

'Сохраняем файл

 

File.Close

MsgBox("Готово!")
