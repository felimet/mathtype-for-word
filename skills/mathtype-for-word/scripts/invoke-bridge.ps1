[CmdletBinding(PositionalBinding = $false)]
param(
    [Parameter(ValueFromRemainingArguments)]
    [string[]]$BridgeArguments
)

# Runs under Windows PowerShell 5.1 or PowerShell 7; the bridge itself always runs in PowerShell 7.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$candidates = @()
if (-not [string]::IsNullOrWhiteSpace($env:MATHTYPE_WORD_PLUGIN_ROOT)) {
    $candidates += Join-Path $env:MATHTYPE_WORD_PLUGIN_ROOT 'scripts\mathtype-word.ps1'
}
$candidates += [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\scripts\mathtype-word.ps1'))
$bridge = $candidates | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
if ($null -eq $bridge) {
    throw 'The full mathtype-for-word plugin is not installed. Register its MCP server or set MATHTYPE_WORD_PLUGIN_ROOT.'
}
$found = & (Join-Path (Split-Path -Parent $bridge) 'ensure-pwsh.ps1') | Select-Object -Last 1 | ConvertFrom-Json
if (-not $found.ok) {
    throw "PowerShell 7 is required, and installing it automatically failed: $($found.error) Install it with 'winget install --id Microsoft.PowerShell --source winget --installer-type wix', then retry."
}
if ($found.installed) {
    [Console]::Error.WriteLine("[NOTICE] PowerShell 7 was not installed, so PowerShell $($found.version) was installed for the current user in $($found.folder) from $($found.source) (SHA-256 checked, signed by Microsoft). Only this plugin uses it; PATH is unchanged. Tell the user.")
}
& $found.path -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $bridge @BridgeArguments
exit $LASTEXITCODE
