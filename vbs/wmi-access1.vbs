On Error Resume Next
Set objService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")
If Err.Number <> 0 Then
	WScript.Echo Err.Number & ": " & Err.Description
	WScript.Quit
End If
i=0
dim text
For Each objProc In objService.ExecQuery("SELECT * FROM Win32_Processor")
i=i+1
Next
text="Раздел " & chr(34) & "Краткая сводка" & chr(34) & vbCrlf
text=text & "Кол-во процессоров - " & i
For Each objPhMem In objService.ExecQuery("SELECT * FROM Win32_PhysicalMemory")
	i = objPhMem.Capacity/1024/1204
Next
text=text & vbCrlf & "Объем оперативной памяти - " & i '& "Мегабайт"
WScript.Echo text