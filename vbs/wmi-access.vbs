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
	File.Write("����� �����������: " & i & VbCrLf) 
	Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory")
	i=0
	For Each objItem in colItems
	i=i+ objItem.Capacity/1024/1024	
	Next
	File.Write("����� ����������� ������: " & i & VbCrLf)
	File.Write("������� ��� ������:" & NamePC & VbCrLf)
	Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem",,48)
	For Each objItem in colItems
	File.Write("������ ��: " & objItem.Caption & " " & objItem.OSArchitecture & VbCrLf) 
	Next
	Set colItems = objWMIService.ExecQuery("Select * from Win32_DiskDrive")
	i=0
	For Each objItem in colItems
		If not objItem.Size = " " Then i=i+ objItem.Size
	Next
	i=i/1024/1024/1024
	File.Write("������ ����� �������� �����: " & i & " ��" & VbCrLf) 
	Set colItems = objWMIService.ExecQuery("Select * from Win32_LogicalDisk")
	i=0
	For Each objItem in colItems
		If not objItem.Size = " " Then i=i+ objItem.FreeSpace
	Next
	i=i/1024/1024/1024
	File.Write("��������� ����� �������� �����: " & i & " ��" & VbCrLf)  
End Sub

Sub ListHardware()
	Set colItems = objWMIService.ExecQuery("Select * from Win32_PnPEntity")

	For Each objItem in colItems
		File.Write("��������: " & objItem.Description & VbCrLf)  
		File.Write("�������������: " & objItem.Manufacturer & VbCrLf) 
		File.Write("���: " & objItem.Name & VbCrLf) 
		File.Write("������: " & objItem.Status & VbCrLf)  
	Next
End Sub

Sub ListSharedFolders
	Set colItems = objWMIService.ExecQuery("Select * from Win32_Share WHERE Type = 0")
	File.Write("������ ������� �����" & VbCrLf)
	For Each objItem in colItems
		File.Write("���: " & objItem.Name & VbCrLf)
		File.Write("����: " & objItem.Path & VbCrLf)
		File.Write("���: " & objItem.Type & VbCrLf)
	Next
	Set colItems = objWMIService.ExecQuery("Select * from Win32_Share WHERE Type = 1")
	File.Write("������ ���������" & VbCrLf)
	For Each objItem in colItems
		File.Write("���: " & objItem.Name & VbCrLf)
		File.Write("����: " & objItem.Path & VbCrLf)
		File.Write("���: " & objItem.Type & VbCrLf)
	Next
End Sub

Sub ListSoftware
	Set colItems = objWMIService.ExecQuery("Select Name, Caption, Vendor, Version from Win32_Product")
	File.Write("������ ��" & VbCrLf)
	For Each objItem in colItems
		File.Write("���: " & objItem.Name & VbCrLf)
		File.Write("��������: " & objItem.Caption & VbCrLf) 
		File.Write("�������������: " & objItem.Vendor & VbCrLf) 
		File.Write("������: " & objItem.Version & VbCrLf) 
	Next
End Sub

'-------------------------------------------------------------------------------



ListShortInfo
' ������������	
ListHardware

'������� �����
ListSharedFolders

'������������� ����������
ListSoftware
File.Close
MsgBox("������!!!")
