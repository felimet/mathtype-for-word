[CmdletBinding()]
param([switch]$KeepArtifacts)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$bridge = Join-Path $root 'scripts\mathtype-word.ps1'
$inputPath = Join-Path $root 'evals\fixtures\en-presentation-draft.pptx'
$manifestPath = Join-Path $root 'evals\fixtures\en-powerpoint-manifest.json'
$testDirectory = Join-Path $env:TEMP ('mathtype-powerpoint-test-' + [Guid]::NewGuid().ToString('N'))
$outputPath = Join-Path $testDirectory 'output.pptx'
$baselineWordPids = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$baselinePowerPointPids = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

try {
    New-Item -ItemType Directory -Path $testDirectory | Out-Null
    $probeOutput = & pwsh.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File $bridge -Action probe-pptx
    $probe = $probeOutput | Select-Object -Last 1 | ConvertFrom-Json
    if (-not $probe.ok) { throw "PowerPoint probe failed: $($probe.error)" }
    if (-not $probe.word_ready) { throw 'probe-pptx requires hidden Word conversion readiness.' }

    $renderOutput = & pwsh.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File $bridge `
        -Action render-pptx -InputPath $inputPath -OutputPath $outputPath -ManifestPath $manifestPath
    $render = $renderOutput | Select-Object -Last 1 | ConvertFrom-Json
    if (-not $render.ok) { throw "PowerPoint render failed: $($render.error)" }

    $validationOutput = & pwsh.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File $bridge `
        -Action validate-pptx -InputPath $outputPath -ManifestPath $manifestPath
    $validation = $validationOutput | Select-Object -Last 1 | ConvertFrom-Json
    if (-not $validation.ok) { throw "PowerPoint validation failed: $($validation.errors -join '; ')" }
    if ($validation.counts.mathtype_objects -ne 1) { throw 'Expected exactly one MathType OLE object.' }
    if ($validation.counts.mathml_verified -ne 1) { throw 'Expected the embedded MathML to match the manifest.' }

    $newWordPids = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $baselineWordPids -notcontains $_.Id })
    if ($newWordPids.Count -gt 0) { throw "PowerPoint integration left Word running: $($newWordPids.Id -join ', ')" }

    [ordered]@{ ok = $true; input = $inputPath; output = $outputPath; probe = $probe; render = $render; validation = $validation; new_word_pids = @() } |
        ConvertTo-Json -Depth 12
}
finally {
    $exitDeadline = [DateTime]::UtcNow.AddSeconds(5)
    do {
        $newPowerPointProcesses = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object { $baselinePowerPointPids -notcontains $_.Id })
        if ($newPowerPointProcesses.Count -eq 0) { break }
        Start-Sleep -Milliseconds 250
    } while ([DateTime]::UtcNow -lt $exitDeadline)
    if ($newPowerPointProcesses.Count -gt 0) {
        [Console]::Error.WriteLine('[WARN] Test-created PowerPoint process(es) remain; not terminating automatically: ' + (($newPowerPointProcesses.Id) -join ', '))
    }
    if (-not $KeepArtifacts -and (Test-Path -LiteralPath $testDirectory)) {
        Remove-Item -LiteralPath $testDirectory -Recurse -Force
    }
    elseif ($KeepArtifacts) { [Console]::Error.WriteLine("[INFO] Kept test artifacts: $testDirectory") }
}
