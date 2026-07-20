[CmdletBinding()]
param([switch]$KeepArtifacts)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$bridge = Join-Path $root 'scripts\mathtype-word.ps1'
$manifest = Join-Path $PSScriptRoot 'fixtures\live-manifest.json'
$testDirectory = Join-Path $env:TEMP ('mathtype-word-test-' + [Guid]::NewGuid().ToString('N'))
$inputPath = Join-Path $testDirectory 'input.docx'
$outputPath = Join-Path $testDirectory 'output.docx'
$baselineWordPids = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

try {
    New-Item -ItemType Directory -Path $testDirectory | Out-Null
    & (Join-Path $PSScriptRoot 'create-live-fixture.ps1') -OutputPath $inputPath
    if (-not (Test-Path -LiteralPath $inputPath)) { throw 'Fixture DOCX was not created.' }

    $renderOutput = & pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $bridge `
        -Action render -InputPath $inputPath -OutputPath $outputPath -ManifestPath $manifest
    $render = $renderOutput | Select-Object -Last 1 | ConvertFrom-Json
    if (-not $render.ok) { throw "Render failed: $($render.error)" }

    $validationOutput = & pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $bridge `
        -Action validate -InputPath $outputPath -ManifestPath $manifest
    $validation = $validationOutput | Select-Object -Last 1 | ConvertFrom-Json
    if (-not $validation.ok) { throw "Validation failed: $($validation.errors -join '; ')" }
    if ($validation.counts.mathtype_objects -ne 3) { throw 'Expected exactly three MathType objects.' }
    if ($validation.counts.native_number_fields -ne 2) { throw 'Expected exactly two native number fields.' }
    if ($validation.counts.native_reference_fields -ne 2) { throw 'Expected exactly two native reference fields.' }
    if (($validation.equation_numbers -join ',') -ne '1,2') { throw 'Expected equation numbers 1,2.' }

    [ordered]@{
        ok = $true
        input = $inputPath
        output = $outputPath
        render = $render
        validation = $validation
    } | ConvertTo-Json -Depth 12
}
finally {
    $newWordProcesses = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $baselineWordPids -notcontains $_.Id })
    if ($newWordProcesses.Count -gt 0) {
        [Console]::Error.WriteLine('[WARN] Test-created Word process(es) remain; not terminating automatically: ' + (($newWordProcesses.Id) -join ', '))
    }
    if (-not $KeepArtifacts -and (Test-Path -LiteralPath $testDirectory)) {
        Remove-Item -LiteralPath $testDirectory -Recurse -Force
    }
    elseif ($KeepArtifacts) {
        [Console]::Error.WriteLine("[INFO] Kept test artifacts: $testDirectory")
    }
}
