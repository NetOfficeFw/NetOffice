@echo off
setlocal EnableDelayedExpansion

set _version=1.8.0.0

mkdir out

xcopy /s /e /y Documentation out\Documentation\
xcopy /s /e /y Examples out\Examples\
xcopy /s /e /y Tutorials out\Tutorials\
xcopy /s /e /y Source out\Source\

nuget restore Source\NetOffice.sln
msbuild Source\NetOffice.sln /t:Build /p:Configuration=Release /v:m

del /s /q Source\ClientApplication\bin\Release\ClientApplication.*
del /s /q Source\ClientApplication\bin\Release\stdole.dll

xcopy /y Source\ClientApplication\bin\Release "out\Assemblies\Any CPU\"

xcopy /y "Breaking Changes.txt" out\
xcopy /y BugFixes.txt out\
xcopy /y ChangeLog.txt out\

pushd out
7z a -tzip ..\NetOffice_v%_version%.zip .
popd
