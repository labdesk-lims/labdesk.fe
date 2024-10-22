Set oShell = WScript.CreateObject("WScript.Shell")
oShell.run "reg delete HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Access\Security /v VBAWarnings /f"