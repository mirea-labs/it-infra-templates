Set objService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")
WScript.Echo 
WScript.Echo "======== System Information ========"
WScript.Echo 
For Each objComSys In objService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
	WScript.Echo "NumberOfprocessors: " & objComSys.NumberOfProcessors
	WScript.Echo "TotalPhysicalMemory: " & objComSys.TotalPhysicalMemory/1024/1024/1024 & " GB"
	WScript.Echo "Name: " & objComSys.Name
next
WScript.Echo 
WScript.Echo "======== OS ========"
WScript.Echo 

For Each objOS In objService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
	WScript.Echo "OS Version: " & objOS.caption
next

WScript.Echo 
WScript.Echo "======== Memmory ========"
WScript.Echo 

For Each objDisc In objService.ExecQuery("SELECT * FROM Win32_DiskDrive")
	WScript.Echo "Polnyi Ob'em ZhestkogoDiska: " & objDisc.Size/1024/1024/1024 & " GB"
next

For Each objLDisc In objService.ExecQuery("SELECT * FROM Win32_LogicalDisk WHERE DriveType = 3")
	WScript.Echo "Svobodnyi Ob'em Logicheskih Discov: " & objLDisc.name & "    " & objLDisc.FreeSpace/1024/1024/1024 & " GB"
next

WScript.Echo 
WScript.Echo "======== OBORUDOVANIE ========"
WScript.Echo 

For Each objPnPE In objService.ExecQuery("SELECT * FROM Win32_PnPEntity")
	WScript.Echo "NameOborudovanie: " & objPnPE.Name
next

WScript.Echo 
WScript.Echo "======== SETEVIE PAPKI & PRINTER ========"
WScript.Echo 

For Each objSh In objService.ExecQuery("SELECT * FROM Win32_Share")
	WScript.Echo "Papki: " & objSh.Path
next

WScript.Echo 
WScript.Echo "======== PO ========"
WScript.Echo 

For Each objPO In objService.ExecQuery("Select Name, Version from Win32_Product")
	WScript.Echo "PO: " & objPO.Name

next



