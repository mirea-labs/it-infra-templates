On Error Resume Next
Set objService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")
i=0
dim text
For Each objProc In objService.ExecQuery("SELECT * FROM Win32_Processor")
i=i+1
Next
text="������ " & chr(34) & "������� ������" & chr(34) & vbCrlf
text=text & "���-�� ����������� - " & i
For Each objPhMem In objService.ExecQuery("SELECT * FROM Win32_PhysicalMemory")
	i = objPhMem.Capacity
Next
text=text & vbCrlf & "����� ����������� ������ - " & i & "��������"
WScript.Echo text