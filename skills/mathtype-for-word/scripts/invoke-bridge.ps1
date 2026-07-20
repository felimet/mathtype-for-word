[CmdletBinding(PositionalBinding = $false)]
param(
    [Parameter(ValueFromRemainingArguments)]
    [string[]]$BridgeArguments
)

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
& pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $bridge @BridgeArguments
exit $LASTEXITCODE
