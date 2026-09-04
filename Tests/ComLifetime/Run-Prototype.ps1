$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$root = Split-Path -Parent $root
$nativeProject = Join-Path $PSScriptRoot 'NativeLifetimeFixture\NativeLifetimeFixture.vcxproj'
$testProject = Join-Path $PSScriptRoot 'NetOffice.ComLifetimePrototype.Tests\NetOffice.ComLifetimePrototype.Tests.csproj'
$msbuild = 'C:\Program Files\Microsoft Visual Studio\18\Enterprise\MSBuild\Current\Bin\MSBuild.exe'

if (-not (Test-Path $msbuild)) {
    throw "Visual Studio 2026 Enterprise MSBuild was not found at $msbuild"
}

& $msbuild $nativeProject -restore -m -p:Configuration=Release -p:Platform=x64
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

& $msbuild $testProject -restore -m -p:Configuration=Release -p:Platform=x64 -p:RestoreLockedMode=false
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

& dotnet test $testProject --no-build --no-restore --configuration Release --filter 'TestCategory=IntegrationTests'
exit $LASTEXITCODE
