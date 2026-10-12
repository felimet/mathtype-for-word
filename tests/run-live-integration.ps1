[CmdletBinding()]
param([switch]$KeepArtifacts)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$bridge = Join-Path $root 'scripts\mathtype-word.ps1'
$batch = Join-Path $root 'scripts\mathtype_batch.py'
$manifest = Join-Path $PSScriptRoot 'fixtures\live-manifest.json'
$testDirectory = Join-Path $env:TEMP ('mathtype-word-test-' + [Guid]::NewGuid().ToString('N'))
$inputPath = Join-Path $testDirectory 'input.docx'
$outputPath = Join-Path $testDirectory 'output.docx'
$batchedOutputPath = Join-Path $testDirectory '分批 輸出.docx'
$baselineWordPids = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$previousJobsDirectory = $env:MATHTYPE_JOBS_DIR
$python = Get-Command python.exe -ErrorAction SilentlyContinue | Select-Object -First 1
if ($null -eq $python) { $python = Get-Command python -ErrorAction Stop | Select-Object -First 1 }

function Assert-Validation {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Label)
    $output = & pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $bridge `
        -Action validate -InputPath $Path -ManifestPath $manifest
    $validation = $output | Select-Object -Last 1 | ConvertFrom-Json
    if (-not $validation.ok) { throw "$Label validation failed: $($validation.errors -join '; ')" }
    if ($validation.counts.mathtype_objects -ne 6) { throw "$Label expected exactly six MathType objects." }
    if ($validation.counts.native_number_fields -ne 3) { throw "$Label expected exactly three native number fields." }
    if ($validation.counts.native_reference_fields -ne 2) { throw "$Label expected exactly two native reference fields." }
    if (($validation.equation_numbers -join ',') -ne '1,2,3') { throw "$Label expected equation numbers 1,2,3." }
    if (@($validation.warnings | Where-Object { $_ -match 'clipped' }).Count -gt 0) { throw "$Label left a clipped equation." }
    return $validation
}

function Get-DocumentXml {
    param([Parameter(Mandatory)][string]$Path)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $reader = [IO.StreamReader]::new($archive.GetEntry('word/document.xml').Open())
        try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
    }
    finally { $archive.Dispose() }
}

try {
    New-Item -ItemType Directory -Path $testDirectory | Out-Null
    $env:MATHTYPE_JOBS_DIR = Join-Path $testDirectory 'jobs'
    & (Join-Path $PSScriptRoot 'create-live-fixture.ps1') -OutputPath $inputPath
    if (-not (Test-Path -LiteralPath $inputPath)) { throw 'Fixture DOCX was not created.' }

    $renderOutput = & pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $bridge `
        -Action render -InputPath $inputPath -OutputPath $outputPath -ManifestPath $manifest
    $render = $renderOutput | Select-Object -Last 1 | ConvertFrom-Json
    if (-not $render.ok) { throw "Render failed: $($render.error)" }
    if (@($render.line_spacing_changed_for) -notcontains 'clipped_fraction') { throw 'The exactly spaced paragraph was not switched to at-least spacing.' }
    # An at-least line inside the document grid still snaps to the grid pitch; the repair must leave the grid.
    $gridOff = '<w:snapToGrid w:val="0"/>'
    if ((Get-DocumentXml -Path $inputPath).Contains($gridOff)) { throw 'The fixture already leaves the document grid.' }
    if (-not (Get-DocumentXml -Path $outputPath).Contains($gridOff)) { throw 'The repaired paragraph still snaps to the document grid.' }
    $validation = Assert-Validation -Path $outputPath -Label 'Single-session render'

    # One equation per Word session: the unreferenced numbered equation is committed in an early chunk, and the
    # referenced ones plus both references come last. Every chunk must validate and the output publish once.
    $batchOutput = & $python.Source $batch render --input $inputPath --output $batchedOutputPath --manifest $manifest --batch-size 1
    $batched = $batchOutput | Select-Object -Last 1 | ConvertFrom-Json
    if (-not $batched.ok) { throw "Batched render failed: $($batched.error)" }
    if ($batched.chunks -lt 4) { throw "Expected the batched render to use several chunks; got $($batched.chunks)." }
    # This caller decodes with the console code page; ASCII JSON must still carry the CJK path intact.
    # Compare the CJK file name: the job runner reports a resolved path, which expands an 8.3 short %TEMP%.
    if ([IO.Path]::GetFileName($batched.output_path) -cne '分批 輸出.docx' -or -not (Test-Path -LiteralPath $batched.output_path)) {
        throw "Batched output_path did not round-trip: $($batched.output_path)"
    }
    if (@($batched.line_spacing_changed_for) -notcontains 'clipped_fraction') { throw 'Batched render lost an earlier chunk''s line-spacing change.' }
    $batchedValidation = Assert-Validation -Path $batchedOutputPath -Label 'Batched render'
    $status = & $python.Source $batch status --output $batchedOutputPath | Select-Object -Last 1 | ConvertFrom-Json
    if ($status.status -ne 'completed') { throw "Batched job status is $($status.status)." }
    if ([IO.Path]::GetFileName($status.output_path) -cne '分批 輸出.docx' -or -not (Test-Path -LiteralPath $status.output_path)) {
        throw "Status output_path did not round-trip: $($status.output_path)"
    }

    [ordered]@{
        ok = $true
        input = $inputPath
        output = $outputPath
        render = $render
        validation = $validation
        batched_chunks = $batched.chunks
        batched_validation = $batchedValidation
    } | ConvertTo-Json -Depth 12
}
finally {
    if ($null -eq $previousJobsDirectory) { Remove-Item Env:\MATHTYPE_JOBS_DIR -ErrorAction SilentlyContinue }
    else { $env:MATHTYPE_JOBS_DIR = $previousJobsDirectory }
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
