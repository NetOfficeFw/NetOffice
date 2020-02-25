regsvr32.exe bin\Debug\netcoreapp3.1\SimpleNetCoreAddin.comhost.dll

reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\NetOfficeSimpleNetCoreAddin.WordAddin" /v LoadBehavior /t REG_DWORD /d 3
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\NetOfficeSimpleNetCoreAddin.WordAddin" /v FriendlyName /t REG_SZ /d "NetOffice Word Addin (.NET Core 3.1)"
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\NetOfficeSimpleNetCoreAddin.WordAddin" /v Description /t REG_SZ /d "Sample addin running in .NET Core 3.1"
