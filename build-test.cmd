@echo off
setlocal EnableDelayedExpansion

msbuild -m Source\NetOffice2.sln /t:Build;Pack /p:Configuration=Release /v:m ^
  /p:ContinuousIntegrationBuild=true ^
  /p:VersionSuffix=preview1 /p:PackageOutputPath="%~dp0out" ^
  /p:RepositoryBranch="preview/netcore3" /p:RepositoryCommit="ed0ca2361e3c4909ca272c950444376d540b71a8" ^
  /p:SignOutput=true
