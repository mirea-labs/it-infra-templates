Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")

i=0
	
dim text
For Each objProc In objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
	i=i+1
Next
text="Краткая сводка:" & vbCrlf & vbCrlf
text=text & "Кол-во процессоров - " & i

i=0
For Each objPhMem In objWMIService.ExecQuery("SELECT * FROM Win32_PhysicalMemory")
	i = i+ objPhMem.Capacity/1024/1024
Next
text = text & vbCrlf & "Объем оперативной памяти - " & i & " Мегабайт"

For Each objMoth In objWMIService.ExecQuery("SELECT * FROM Win32_MotherboardDevice")
	text = text & vbCrlf & "Сетевое имя машины - " & objMoth.SystemName
	PC_Name = objMoth.SystemName
Next

For Each objItem in objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem",,48) 
	text = text & vbCrlf & "Версия ОС - " & objItem.Caption & " " & objItem.OSArchitecture
Next

i=0
For Each objDisk In objWMIService.ExecQuery("SELECT * FROM Win32_DiskDrive")
	if Not objDisk.Size = "" Then i = i + objDisk.Size
Next
i = i/1024/1024/1024
text = text & vbCrlf & "Полный объем жестких дисков - " & i & " Гигабайт"

i=0
For Each objDisk In objWMIService.ExecQuery ("Select * From Win32_LogicalDisk")
	if Not objDisk.Size = "" Then i = i + objDisk.FreeSpace
Next
i = i/1024/1024/1024
text = text & vbCrlf & "Свободный объем жестких дисков - " & i & " Гигабайт" & vbCrlf & "###########################################################################################################" & vbCrlf & "Оборудование:" & vbCrlf

For Each objItem in objWMIService.ExecQuery("Select * from Win32_PnPEntity")
	text = text & vbCrlf & "Описание: " & objItem.Description
	text = text & vbCrlf & "Производитель: " & objItem.Manufacturer
	text = text & vbCrlf & "Имя: " & objItem.Name
	text = text & vbCrlf & "Статус: " & objItem.Status & vbCrlf
Next
text = text & "###########################################################################################################" & vbCrlf & "Ресурсы:" & vbCrlf

For Each objItem in objWMIService.ExecQuery("Select * from Win32_Share where type<>1")
	text = text & vbCrlf & objItem.Caption & " " & objItem.Name
Next
text = text & vbCrlf
For Each objItem in objWMIService.ExecQuery("Select * from Win32_Share where type=1")
	text = text & vbCrlf & objItem.Caption & " " & objItem.Name
Next

text = text & "###########################################################################################################" & vbCrlf & "Программное обеспечение:" & vbCrlf
For Each objItem in objWMIService.ExecQuery("Select * from Win32_Product")
	text = text & vbCrlf & objItem.Name & objItem.Version
Next


'Сохраняем файл
set FSO = CreateObject("Scripting.FileSystemObject")
If FSO.FileExists(PC_Name & ".txt") Then 
	Set file = FSO.GetFile(PC_Name & ".txt")
	File.Delete
End If
set File = FSO.OpenTextFile(PC_Name & ".txt", 8, True)
 
File.Write(text)
File.Close

MsgBox("Готово!")
