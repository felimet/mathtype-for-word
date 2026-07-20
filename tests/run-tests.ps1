[CmdletBinding()]
param([Alias('IncludeLiveWord')][switch]$IncludeLiveOffice, [switch]$KeepArtifacts)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$python = Get-Command python.exe -ErrorAction SilentlyContinue | Select-Object -First 1
if ($null -eq $python) { $python = Get-Command python -ErrorAction Stop | Select-Object -First 1 }

$previousLiveSetting = $env:MATHTYPE_OFFICE_LIVE_TEST
$previousBytecodeSetting = $env:PYTHONDONTWRITEBYTECODE
if ($IncludeLiveOffice) { $env:MATHTYPE_OFFICE_LIVE_TEST = '1' }
$env:PYTHONDONTWRITEBYTECODE = '1'
& $python.Source -m unittest discover -s $PSScriptRoot -p 'test_*.py' -v
$pythonExitCode = $LASTEXITCODE
if ($null -eq $previousLiveSetting) { Remove-Item Env:\MATHTYPE_OFFICE_LIVE_TEST -ErrorAction SilentlyContinue }
else { $env:MATHTYPE_OFFICE_LIVE_TEST = $previousLiveSetting }
if ($null -eq $previousBytecodeSetting) { Remove-Item Env:\PYTHONDONTWRITEBYTECODE -ErrorAction SilentlyContinue }
else { $env:PYTHONDONTWRITEBYTECODE = $previousBytecodeSetting }
if ($pythonExitCode -ne 0) { throw "Python tests failed with exit code $pythonExitCode." }

$parseErrors = $null
$tokens = $null
foreach ($scriptPath in Get-ChildItem -LiteralPath $root -Recurse -Filter '*.ps1') {
    $null = [Management.Automation.Language.Parser]::ParseFile($scriptPath.FullName, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count -gt 0) {
        throw "PowerShell parse failure in $($scriptPath.FullName): $($parseErrors.Message -join '; ')"
    }
}

if ($IncludeLiveOffice) {
    & (Join-Path $PSScriptRoot 'run-live-integration.ps1') -KeepArtifacts:$KeepArtifacts
    & (Join-Path $PSScriptRoot 'run-live-powerpoint-integration.ps1') -KeepArtifacts:$KeepArtifacts
}

[ordered]@{ ok = $true; python_tests = 'passed'; powershell_parse = 'passed'; live_office = [bool]$IncludeLiveOffice } |
    ConvertTo-Json -Compress
