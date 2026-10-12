# Starts the MCP server for registrations made before 1.4.0, which run this file; current registrations start
# python scripts\mcp_server.py directly. Runs under Windows PowerShell 5.1 or PowerShell 7.
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$server = Join-Path $PSScriptRoot 'mcp_server.py'
$python = Get-Command python.exe -ErrorAction SilentlyContinue | Select-Object -First 1
if ($null -eq $python) {
    $python = Get-Command python -ErrorAction SilentlyContinue | Select-Object -First 1
}
if ($null -eq $python) {
    [Console]::Error.WriteLine('[ERROR] Python 3 was not found on PATH.')
    exit 1
}
& $python.Source $server
exit $LASTEXITCODE
