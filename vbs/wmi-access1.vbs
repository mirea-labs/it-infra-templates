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
Print "Êðàòêàÿ ñâîäêà:" & vbCrlf & vbCrlf
Print "Êîë-âî ïðîöåññîðîâ - " & i

i=0
For Each objPhMem In objWMIService.ExecQuery("SELECT * FROM Win32_PhysicalMemory")
	i = i+ objPhMem.Capacity/1024/1024
Next
Print vbCrlf & "Îáúåì îïåðàòèâíîé ïàìÿòè - " & i & " Ìåãàáàéò"

Print vbCrlf & "Ñåòåâîå èìÿ ìàøèíû - " & PC_Name

For Each objItem in objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem",,48) 
	Print vbCrlf & "Âåðñèÿ ÎÑ - " & objItem.Caption & " " & objItem.OSArchitecture
Next

i=0
For Each objDisk In objWMIService.ExecQuery("SELECT * FROM Win32_DiskDrive")
	if Not objDisk.Size = "" Then i = i + objDisk.Size
Next
i = i/1024/1024/1024
Print vbCrlf & "Ïîëíûé îáúåì æåñòêèõ äèñêîâ - " & i & " Ãèãàáàéò"

i=0
For Each objDisk In objWMIService.ExecQuery ("Select * From Win32_LogicalDisk")
	if Not objDisk.Size = "" Then i = i + objDisk.FreeSpace
Next
i = i/1024/1024/1024
Print vbCrlf & "Ñâîáîäíûé îáúåì æåñòêèõ äèñêîâ - " & i & " Ãèãàáàéò" & vbCrlf & "###########################################################################################################" & vbCrlf & "Îáîðóäîâàíèå:" & vbCrlf

For Each objItem in objWMIService.ExecQuery("Select * from Win32_PnPEntity")
	Print vbCrlf & "Îïèñàíèå: " & objItem.Description
	Print vbCrlf & "Ïðîèçâîäèòåëü: " & objItem.Manufacturer
	Print vbCrlf & "Èìÿ: " & objItem.Name
	Print vbCrlf & "Ñòàòóñ: " & objItem.Status & vbCrlf
Next
Print "###########################################################################################################" & vbCrlf & "Ðåñóðñû:" & vbCrlf

For Each objItem in objWMIService.ExecQuery("Select * from Win32_Share where type<>1")
	Print vbCrlf & objItem.Caption & " " & objItem.Name
Next
Print vbCrlf
For Each objItem in objWMIService.ExecQuery("Select * from Win32_Share where type=1")
	Print vbCrlf & objItem.Caption & " " & objItem.Name
Next

Print "###########################################################################################################" & vbCrlf & "Ïðîãðàììíîå îáåñïå÷åíèå:" & vbCrlf
For Each objItem in objWMIService.ExecQuery("Select * from Win32_Product")
	Print vbCrlf & objItem.Name & objItem.Version
Next

Sub Print(text)
	'WScript.Echo text
	File.Write(text)
End Sub

'Ñîõðàíÿåì ôàéë

 

File.Close

MsgBox("Ãîòîâî!")