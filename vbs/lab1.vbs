Set objService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")
i=0
For Each objProc In objService.ExecQuery("SELECT * FROM Win32_Processor")
	i=i+1
Next
Wscript.echo "���������� �����������: " & i
For Each objphmem In objService.ExecQuery("SELECT * FROM Win32_PhysicalMemory")
	WScript.Echo "����� ����������� ������: " & ((objPhMem.Capacity)/1024/1024) & " �����"
Next
d=0
For Each objNtw In objService.ExecQuery("SELECT * FROM Win32_NetworkAdapter")
	WScript.Echo "������� ��� ���������: " & objNtw.SystemName
	d=d+1
	if d=1 then exit for
Next
For Each objWin In objService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
	WScript.Echo "������ ��: " & objWin.caption & " " & objWin.OSArchitecture
Next
Wscript.Echo ""
For Each objDisk In objService.ExecQuery("SELECT * FROM Win32_LogicalDisk where drivetype=3")
	Wscript.Echo "�������� �������:" & " " & objDisk.caption
	WScript.Echo "������ ����� �������: " & ((objDisk.Size)/1024/1024/1024) & " �����"
	WScript.Echo "��������� ����� �������: " & ((objDisk.FreeSpace)/1024/1024/1024) & " �����"
Next
Wscript.Echo ""
Wscript.Echo "�������� ������������:"
For Each objHard In objService.ExecQuery("SELECT * FROM Win32_PNPEntity")
	WScript.Echo objHard.Name
Next
Wscript.Echo ""
Wscript.echo "�������� ������� ����� ������:"
For Each objShare In objService.ExecQuery("SELECT * FROM Win32_Share where type=0")
	WScript.Echo objShare.Name
Next
Wscript.Echo ""
Wscript.echo "������ ���������:"
For Each objPrinter In objService.ExecQuery("SELECT * FROM Win32_Printer")
	WScript.Echo objPrinter.Name
Next
Wscript.Echo ""
Wscript.echo "������ �������������� ��:"
For Each objSoft In objService.ExecQuery("SELECT * FROM Win32_Product")
	WScript.Echo objSoft.Name
Next