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
        'The second native reference is {{EQREF:quadratic}}.',
        'An unreferenced numbered display follows:',
        '{{MATH:energy_flux}}',
        '其中，q 為熱通量，單位為 W/m²；k 為熱傳導係數；T 為溫度。'
    )
    foreach ($paragraph in $paragraphs) {
        $selection.TypeText($paragraph)
        $selection.TypeParagraph()
    }
    # Exactly spaced text clips a tall inline equation unless the bridge switches it to at-least spacing.
    $selection.ParagraphFormat.LineSpacingRule = 4
    $selection.ParagraphFormat.LineSpacing = 12
    $selection.TypeText('Exact 12 pt spacing holds {{MATH:clipped_fraction}} inline.')
    $selection.TypeParagraph()
    $selection.ParagraphFormat.LineSpacingRule = 0
    # Word's Find can return a table-cell match again; the marker search must still terminate.
    $table = $document.Tables.Add($selection.Range, 1, 2)
    $table.Cell(1, 1).Range.Text = 'Cell value'
    $table.Cell(1, 2).Range.Text = 'Ratio {{MATH:table_cell}}'
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($table)
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
