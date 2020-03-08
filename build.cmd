@echo off
setlocal EnableDelayedExpansion

set _version=1.8.0
set _suffix=alpha01
::  goIT Solutions, s.r.o. certificate
set _thumbprint=ac6dbffb1bf8b62281deb8641023a66cddc5db57

mkdir out

xcopy /s /e /y Documentation out\Documentation\
xcopy /s /e /y Examples out\Examples\
xcopy /s /e /y Tutorials out\Tutorials\
xcopy /s /e /y Source out\Source\

msbuild Source\NetOffice.sln /t:Restore;Build;Pack /p:IncludeSource=true /p:Configuration=Release /p:VersionSuffix="%_suffix%" /p:PackageOutputPath="%CD%\out" /v:m
nuget.exe sign "out\*.nupkg" -CertificateFingerprint "%_thumbprint%" -HashAlgorithm SHA256 -Timestamper http://timestamp.comodoca.com -TimestampHashAlgorithm SHA256 -Overwrite -OutputDirectory out -NonInteractive -ForceEnglishOutput

del /s /q Source\ClientApplication\bin\Release\ClientApplication.*
del /s /q Source\ClientApplication\bin\Release\stdole.dll

xcopy /y Source\ClientApplication\bin\Release "out\Assemblies\Any CPU\"

rem del /s /q Source\ClientApplication\bin\Release\ClientApplication.*
rem del /s /q Source\ClientApplication\bin\Release\*\stdole.dll

xcopy /y /s /i /e Source\ClientApplication\bin\Release "out\Assemblies\"

xcopy /y "Breaking Changes.txt" out\
xcopy /y BugFixes.txt out\
xcopy /y ChangeLog.txt out\

pushd out
7z a -tzip ..\NetOffice_v%_version%-%_suffix%.zip .
popd
