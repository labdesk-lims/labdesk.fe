Set oShell = WScript.CreateObject("WScript.Shell")
oShell.run "reg add HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Access\Security /v VBAWarnings /t REG_DWORD /d 00000000 /f"