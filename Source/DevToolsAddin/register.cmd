C:\Windows\SysWOW64\regsvr32.exe /s DevToolsAddin.comhost.dll
::C:\Windows\System32\regsvr32.exe /s DevToolsAddin.comhost.dll

reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\NetOffice.DevToolsAddin" /f /v LoadBehavior /t REG_DWORD /d 3
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\NetOffice.DevToolsAddin" /f /v FriendlyName /t REG_SZ /d "NetOffice DevTools for PowerPoint"
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\NetOffice.DevToolsAddin" /f /v Description /t REG_SZ /d "Run playwright tests for Microsoft Office applications."
