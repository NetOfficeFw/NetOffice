@echo off
setlocal EnableDelayedExpansion

set _version=1.7.4.9
set _certificate=goIT Solutions, s.r.o.
set _thumbprint=AC6DBFFB1BF8B62281DEB8641023A66CDDC5DB57

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
signtool.exe sign /v /fd sha256 /td sha256 /sha1 "%_thumbprint%" /tr http://timestamp.comodoca.com/rfc3161 "out\Assemblies\Any CPU\*.dll"

xcopy /y "Breaking Changes.txt" out\
xcopy /y BugFixes.txt out\
xcopy /y ChangeLog.txt out\
xcopy /y LICENSE.txt out\

pushd out
7z a -tzip ..\NetOffice_v%_version%.zip .
popd
