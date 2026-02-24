Set objShell = CreateObject("WScript.Shell")
objShell.Run "PowerShell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & Replace(WScript.ScriptFullName, ".vbs", ".ps1") & """", 0, False
