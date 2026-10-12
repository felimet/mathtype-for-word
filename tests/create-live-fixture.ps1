[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$OutputPath,
    # A draft with plain-text mathematics and no markers, for scan and mark.
    [switch]$PlainText
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
    if ($PlainText) {
        # Word sub/superscripts that plain text loses: in the first two paragraphs, at the end of a table cell, and
        # after the table. Literal carets test mark's escaping for Word Find.
        $selection.TypeText('Plain-text drafts type formulas such as E = mc')
        $selection.Font.Superscript = -1
        $selection.TypeText('2')
        $selection.Font.Superscript = 0
        $selection.TypeText(' inside sentences.')
        $selection.TypeParagraph()
        $selection.TypeText('x = (' + [char]0x2212 + 'b ' + [char]0x00B1 + ' ' + [char]0x221A + '(b')
        $selection.Font.Superscript = -1
        $selection.TypeText('2')
        $selection.Font.Superscript = 0
        $selection.TypeText(' ' + [char]0x2212 + ' 4ac)) / 2a')
        $selection.TypeParagraph()
        $selection.TypeText('As shown in Eq. (1), the roots follow from the coefficients.')
        $selection.TypeParagraph()
        $selection.TypeText('Pythagoras: a^2 + b^2 = c^2')
        $selection.TypeParagraph()
        $selection.TypeText('Water is H')
        $selection.Font.Subscript = -1
        $selection.TypeText('2')
        $selection.Font.Subscript = 0
        $selection.TypeText('O.')
        $selection.TypeParagraph()
        $table = $document.Tables.Add($selection.Range, 1, 2)
        $table.Cell(1, 1).Range.Text = [string][char]0x03B1 + ' + ' + [char]0x03B2
        $cell = $table.Cell(1, 2).Range
        $cell.Text = 'area m2'
        $squared = $document.Range($cell.Start + 6, $cell.Start + 7) # the 2 of 'area m2'
        $squared.Font.Superscript = -1
        foreach ($value in @($squared, $cell, $table)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value) }
        $document.SaveAs2([IO.Path]::GetFullPath($OutputPath), 16)
        return
    }
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
    $table = $document.Tables.Add($selection.Range, 2, 2)
    $table.Cell(1, 2).Range.Text = 'Ratio {{MATH:table_cell}}'
    $table.Cell(2, 2).Range.Text = 'Second row'
    # An exact row height clips the fraction unless the bridge switches the row to an at-least height. The
    # vertically merged first column makes Word refuse Rows.Item for the whole table.
    $row = $table.Rows.Item(1)
    $row.HeightRule = 2
    $row.Height = 14
    $table.Cell(1, 1).Merge($table.Cell(2, 1))
    $table.Cell(1, 1).Range.Text = 'Merged label'
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($row)
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
