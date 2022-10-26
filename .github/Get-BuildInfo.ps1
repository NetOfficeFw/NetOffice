#Requires -Version 7.0
#Requires -PSEdition Core
#
# Generates deployment information for the build.
#

param (
  [String]
  [Parameter()]
  $ref,
  [String]
  [Parameter()]
  $event_name,
  [String]
  [Parameter()]
  $configuration
)

function Write-GitHubVariable {
  param ($name, $value)
  Write-Output "$name=$value" >> $env:GITHUB_OUTPUT
  Write-Host "  steps.${env:GITHUB_ACTION}.outputs.$name=$value" -ForegroundColor Cyan
}

[xml]$project = Get-Content (Join-Path -Path $PSScriptRoot -ChildPath '..\Source\NetOffice.props')
$app_version = $project.Project.PropertyGroup[0].NetOfficeRelease

$sign_binaries = 'false'
$publish_nuget = 'false'
$app_version_suffix = "preview${env:GITHUB_RUN_NUMBER}"

if ($configuration -ieq 'release') {
  if ($event_name -notlike 'pull_request') {
    $sign_binaries = 'true'
  }

  if ($ref -like 'refs/tags/v*') {
    $publish_nuget = 'true'
    $app_version_suffix = ''
  }

  if ($ref -like 'refs/heads/releases/*') {
    $publish_nuget = 'true'
  }

  # if ($ref -like 'refs/heads/dev/github_actions') {
  #  $publish_nuget = 'true'
  # }
}

$app_version_full = $app_version
if ($app_version_suffix -ne '') {
  $app_version_full += '-' + $app_version_suffix
}

Write-GitHubVariable "app_version" $app_version
Write-GitHubVariable "app_version_suffix" $app_version_suffix
Write-GitHubVariable "app_version_full" $app_version_full
Write-GitHubVariable "sign_binaries" $sign_binaries
Write-GitHubVariable "publish_nuget" $publish_nuget
