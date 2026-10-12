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
    if (@($render.row_height_changed_for) -notcontains 'table_cell') { throw 'The exact table row was not switched to an at-least height.' }
    $exactRow = 'w:hRule="exact"'
    if (-not (Get-DocumentXml -Path $inputPath).Contains($exactRow)) { throw 'The fixture lacks an exact table row height.' }
    if ((Get-DocumentXml -Path $outputPath).Contains($exactRow)) { throw 'The table row holding the equation still has an exact height.' }
    $validation = Assert-Validation -Path $outputPath -Label 'Single-session render'

    # One operation per Word session: every equation in its own session, then each reference in a later one that
    # reaches its target only through the anchor bookmark. Every chunk must validate and the output publish once.
    $batchOutput = & $python.Source $batch render --input $inputPath --output $batchedOutputPath --manifest $manifest --batch-size 1
    $batched = $batchOutput | Select-Object -Last 1 | ConvertFrom-Json
    if (-not $batched.ok) { throw "Batched render failed: $($batched.error)" }
    if ($batched.chunks -ne 8) { throw "Expected one chunk per equation and reference (8); got $($batched.chunks)." }
    # This caller decodes with the console code page; ASCII JSON must still carry the CJK path intact.
    # Compare the CJK file name: the job runner reports a resolved path, which expands an 8.3 short %TEMP%.
    if ([IO.Path]::GetFileName($batched.output_path) -cne '分批 輸出.docx' -or -not (Test-Path -LiteralPath $batched.output_path)) {
        throw "Batched output_path did not round-trip: $($batched.output_path)"
    }
    if (@($batched.line_spacing_changed_for) -notcontains 'clipped_fraction') { throw 'Batched render lost an earlier chunk''s line-spacing change.' }
    if (@($batched.row_height_changed_for) -notcontains 'table_cell') { throw 'Batched render lost an earlier chunk''s row-height change.' }
    $batchedValidation = Assert-Validation -Path $batchedOutputPath -Label 'Batched render'
    $status = & $python.Source $batch status --output $batchedOutputPath | Select-Object -Last 1 | ConvertFrom-Json
    if ($status.status -ne 'completed') { throw "Batched job status is $($status.status)." }
    if ([IO.Path]::GetFileName($status.output_path) -cne '分批 輸出.docx' -or -not (Test-Path -LiteralPath $status.output_path)) {
        throw "Status output_path did not round-trip: $($status.output_path)"
    }

    # A draft with plain-text mathematics: scan it, mark what the scan found, then render the marked copy.
    $plainInput = Join-Path $testDirectory 'plain.docx'
    $candidatesPath = Join-Path $testDirectory 'candidates.json'
    $marksPath = Join-Path $testDirectory 'marks.json'
    $markedPath = Join-Path $testDirectory 'marked.docx'
    $plainOutput = Join-Path $testDirectory 'plain-output.docx'
    & (Join-Path $PSScriptRoot 'create-live-fixture.ps1') -OutputPath $plainInput -PlainText
    $scan = & pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $bridge `
        -Action scan -InputPath $plainInput -OutputPath $candidatesPath | Select-Object -Last 1 | ConvertFrom-Json
    if (-not $scan.ok) { throw "Scan failed: $($scan.error)" }
    $candidates = Get-Content -Raw -LiteralPath $candidatesPath | ConvertFrom-Json
    function Find-Candidate([string]$Text) {
        $match = @($candidates.paragraphs | Where-Object { $_.text.Contains($Text) })
        if ($match.Count -ne 1) { throw "Scan listed $($match.Count) paragraph(s) containing '$Text'." }
        return $match[0]
    }
    $energy = Find-Candidate 'E = mc2'
    $roots = Find-Candidate '4ac'
    $reference = Find-Candidate 'Eq. (1)'
    $pythagoras = Find-Candidate 'a^2'
    $water = Find-Candidate 'H2O'
    $sum = Find-Candidate ([string][char]0x03B1 + ' + ' + [char]0x03B2)
    $area = Find-Candidate 'area m2'
    if ($energy.formatted -isnot [array] -or $reference.formatted -isnot [array]) { throw 'Scan lists must stay JSON arrays.' }
    if (@($candidates.formatted_incomplete).Count -ne 0) { throw "Scan could not list every script: $($candidates.formatted_incomplete -join ',')" }
    $offset = $energy.text.IndexOf('mc2') + 2
    if (@($energy.formatted | Where-Object { $_.kind -eq 'superscript' -and $_.text -eq '2' -and $_.offset -eq $offset }).Count -ne 1) { throw 'Scan missed the Word superscript or its offset.' }
    if (@($area.formatted | Where-Object { $_.kind -eq 'superscript' -and $_.offset -eq 6 }).Count -ne 1) { throw 'Scan missed the superscript at the end of a table cell.' }
    if (@($water.formatted | Where-Object { $_.kind -eq 'subscript' -and $_.text -eq '2' }).Count -ne 1) { throw 'Scan missed the subscript after the table.' }
    if (@($reference.signals | Where-Object kind -eq 'reference').Count -ne 1) { throw 'Scan missed the typed equation reference.' }
    if (-not $sum.in_table -or @($sum.signals | Where-Object kind -eq 'greek').Count -ne 2) { throw 'Scan missed the table cell.' }
    @{
        schema_version = 1
        marks = @(
            @{ paragraph = $energy.paragraph; text = 'E = mc2'; marker = '{{MATH:energy}}' }
            @{ paragraph = $roots.paragraph; text = $roots.text; marker = '{{MATH:roots}}' }
            @{ paragraph = $reference.paragraph; text = '(1)'; marker = '{{EQREF:roots}}' } # Eq. stays in the prose
            @{ paragraph = $pythagoras.paragraph; text = 'a^2 + b^2 = c^2'; marker = '{{MATH:pythagoras}}' }
            @{ paragraph = $sum.paragraph; text = $sum.text; marker = '{{MATH:sum}}' }
        )
    } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $marksPath -Encoding utf8
    $mark = & pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $bridge `
        -Action mark -InputPath $plainInput -OutputPath $markedPath -MarksPath $marksPath | Select-Object -Last 1 | ConvertFrom-Json
    if (-not $mark.ok -or $mark.marks -ne 5) { throw "Mark failed: $($mark.error)" }
    if (-not ((Get-DocumentXml -Path $markedPath) -replace '<[^>]+>', '').Contains('As shown in Eq. {{EQREF:roots}}')) { throw 'Marking the reference changed its prose.' }
    $plainManifest = Join-Path $PSScriptRoot 'fixtures\live-plain-manifest.json'
    $plainRender = & pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $bridge `
        -Action render -InputPath $markedPath -OutputPath $plainOutput -ManifestPath $plainManifest | Select-Object -Last 1 | ConvertFrom-Json
    if (-not $plainRender.ok) { throw "Render of the marked draft failed: $($plainRender.error)" }
    $counts = $plainRender.validation.counts
    if ($counts.mathtype_objects -ne 4 -or $counts.native_number_fields -ne 1 -or $counts.native_reference_fields -ne 1) {
        throw "Unexpected counts for the marked draft: $($counts | ConvertTo-Json -Compress)"
    }

    [ordered]@{
        ok = $true
        input = $inputPath
        output = $outputPath
        render = $render
        validation = $validation
        batched_chunks = $batched.chunks
        batched_validation = $batchedValidation
        plain_scan = $scan
        plain_render_counts = $plainRender.validation.counts
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
