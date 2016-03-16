Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")
i=0
'Оборудование
dim text
For Each objProc In objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
	i=i+1
Next
text="Краткая сводка:" & vbCrlf & vbCrlf
text=text & "Число процессоров - " & i
'Сетевые папки
i=0
For Each objPhMem In objWMIService.ExecQuery("SELECT * FROM Win32_PhysicalMemory")
	i = i+ objPhMem.Capacity/1024/1024
Next
text = text & vbCrlf & "Обьем оперативной памяти - " & i & " Мегабайт"
'Установленные приложения
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
text = text & vbCrlf & "Полный обьем жесткого диска - " & i & " Гигабайт"

i=0
For Each objDisk In objWMIService.ExecQuery ("Select * From Win32_LogicalDisk")
	if Not objDisk.Size = "" Then i = i + objDisk.FreeSpace
Next
i = i/1024/1024/1024
text = text & vbCrlf & "Свободный обьем жесткого диска - " & i & " Гигабайт" & vbCrlf & "" & vbCrlf & "Оборудование:" & vbCrlf

For Each objItem in objWMIService.ExecQuery("Select * from Win32_PnPEntity")
	text = text & vbCrlf & "Description: " & objItem.Description
	text = text & vbCrlf & "Manufacturer: " & objItem.Manufacturer
    text = text & vbCrlf & "Name: " & objItem.Name
	text = text & vbCrlf & "Status: " & objItem.Status & vbCrlf
Next
text = text & vbCrlf & "Ресурсы:" & vbCrlf

For Each objItem in objWMIService.ExecQuery("Select * from Win32_Share where type<>1")
    text = text & vbCrlf & objItem.Caption & objItem.Name
Next
text = text & vbCrlf
For Each objItem in objWMIService.ExecQuery("Select * from Win32_Share where type=1")
	text = text & vbCrlf & objItem.Caption & " " & objItem.Name
Next

    text = text & vbCrlf & "Программное обеспечение:" & vbCrlf
For Each objItem in objWMIService.ExecQuery("Select * from Win32_Product")
	text = text & vbCrlf & objItem.Name & objItem.Version
Next

set FSO = CreateObject("Scripting.FileSystemObject")
If FSO.FileExists(PC_Name & ".txt") Then 
	Set file = FSO.GetFile(PC_Name & ".txt")
	File.Delete
End If
set File = FSO.OpenTextFile(PC_Name & ".txt", 8, True)

File.Write(text)
File.Close

MsgBox("ВЫПОЛНЕНО!")

