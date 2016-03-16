Set objService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")
i=0
For Each objProc In objService.ExecQuery("SELECT * FROM Win32_Processor")
	i=i+1
Next
Wscript.echo "Количество процессоров: " & i
For Each objphmem In objService.ExecQuery("SELECT * FROM Win32_PhysicalMemory")
	WScript.Echo "Объем оперативной памяти: " & ((objPhMem.Capacity)/1024/1024) & " мбайт"
Next
d=0
For Each objNtw In objService.ExecQuery("SELECT * FROM Win32_NetworkAdapter")
	WScript.Echo "Сетевое имя комьютера: " & objNtw.SystemName
	d=d+1
	if d=1 then exit for
Next
For Each objWin In objService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
	WScript.Echo "Версия ОС: " & objWin.caption & " " & objWin.OSArchitecture
Next
Wscript.Echo ""
For Each objDisk In objService.ExecQuery("SELECT * FROM Win32_LogicalDisk where drivetype=3")
	Wscript.Echo "Название раздела:" & " " & objDisk.caption
	WScript.Echo "Полный объем раздела: " & ((objDisk.Size)/1024/1024/1024) & " гбайт"
	WScript.Echo "Свободный объем раздела: " & ((objDisk.FreeSpace)/1024/1024/1024) & " гбайт"
Next
Wscript.Echo ""
Wscript.Echo "Перечень оборудования:"
For Each objHard In objService.ExecQuery("SELECT * FROM Win32_PNPEntity")
	WScript.Echo objHard.Name
Next
Wscript.Echo ""
Wscript.echo "Перечень сетевых папок машины:"
For Each objShare In objService.ExecQuery("SELECT * FROM Win32_Share where type=0")
	WScript.Echo objShare.Name
Next
Wscript.Echo ""
Wscript.echo "Список принтеров:"
For Each objPrinter In objService.ExecQuery("SELECT * FROM Win32_Printer")
	WScript.Echo objPrinter.Name
Next
Wscript.Echo ""
Wscript.echo "Список установленного ПО:"
For Each objSoft In objService.ExecQuery("SELECT * FROM Win32_Product")
	WScript.Echo objSoft.Name
Next