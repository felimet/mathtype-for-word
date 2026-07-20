[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$OutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$word = $null
$document = $null
$selection = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Add()
    $selection = $word.Selection
    $paragraphs = @(
        'The mass-energy relation is:',
        '{{MATH:mass_energy}}',
        '其中，E 為能量，單位為 J；m 為質量，單位為 kg；c 為真空光速，單位為 m/s。',
        'The variance {{MATH:inline_variance}} is an inline quantity.',
        'As shown in {{EQREF:mass_energy}}, energy is proportional to mass.',
        'The quadratic solution is:',
        '{{MATH:quadratic}}',
        '其中，x 為未知數，a、b 與 c 為多項式係數，且 a 不等於 0。',
        'The second native reference is {{EQREF:quadratic}}.'
    )
    foreach ($paragraph in $paragraphs) {
        $selection.TypeText($paragraph)
        $selection.TypeParagraph()
    }
    $document.SaveAs2([IO.Path]::GetFullPath($OutputPath), 16)
}
finally {
    if ($null -ne $document) { try { $document.Close(0) } catch {} }
    if ($null -ne $word) { try { $word.Quit() } catch {} }
    foreach ($value in @($selection, $document, $word)) {
        if ($null -ne $value) {
            try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value) } catch {}
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
