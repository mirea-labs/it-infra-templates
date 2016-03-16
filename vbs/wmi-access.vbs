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
Print "������� ������:" & vbCrlf & vbCrlf
Print "���-�� ����������� - " & i

i=0
For Each objPhMem In objWMIService.ExecQuery("SELECT * FROM Win32_PhysicalMemory")
	i = i+ objPhMem.Capacity/1024/1024
Next
Print vbCrlf & "����� ����������� ������ - " & i & " ��������"

Print vbCrlf & "������� ��� ������ - " & PC_Name

For Each objItem in objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem",,48) 
	Print vbCrlf & "������ �� - " & objItem.Caption & " " & objItem.OSArchitecture
Next

i=0
For Each objDisk In objWMIService.ExecQuery("SELECT * FROM Win32_DiskDrive")
	if Not objDisk.Size = "" Then i = i + objDisk.Size
Next
i = i/1024/1024/1024
Print vbCrlf & "������ ����� ������� ������ - " & i & " ��������"

i=0
For Each objDisk In objWMIService.ExecQuery ("Select * From Win32_LogicalDisk")
	if Not objDisk.Size = "" Then i = i + objDisk.FreeSpace
Next
i = i/1024/1024/1024
Print vbCrlf & "��������� ����� ������� ������ - " & i & " ��������" & vbCrlf & "###########################################################################################################" & vbCrlf & "������������:" & vbCrlf

For Each objItem in objWMIService.ExecQuery("Select * from Win32_PnPEntity")
	Print vbCrlf & "��������: " & objItem.Description
	Print vbCrlf & "�������������: " & objItem.Manufacturer
	Print vbCrlf & "���: " & objItem.Name
	Print vbCrlf & "������: " & objItem.Status & vbCrlf
Next
Print "###########################################################################################################" & vbCrlf & "�������:" & vbCrlf

For Each objItem in objWMIService.ExecQuery("Select * from Win32_Share where type<>1")
	Print vbCrlf & objItem.Caption & " " & objItem.Name
Next
Print vbCrlf
For Each objItem in objWMIService.ExecQuery("Select * from Win32_Share where type=1")
	Print vbCrlf & objItem.Caption & " " & objItem.Name
Next

Print "###########################################################################################################" & vbCrlf & "����������� �����������:" & vbCrlf
For Each objItem in objWMIService.ExecQuery("Select * from Win32_Product")
	Print vbCrlf & objItem.Name & objItem.Version
Next

Sub Print(text)
	'WScript.Echo text
	File.Write(text)
End Sub

'��������� ����

 

File.Close

MsgBox("������!")
