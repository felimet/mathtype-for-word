[CmdletBinding()]
param(
    [string]$OutputDirectory = (Join-Path (Split-Path -Parent $PSScriptRoot) 'evals\fixtures'),
    [switch]$OnlyAutoClassification
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Log {
    param([Parameter(Mandatory)][string]$Message)
    [Console]::Error.WriteLine("[fixture] $Message")
}

function Release-ComObject {
    param($Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) }
        catch { Write-Log "COM release warning: $($_.Exception.Message)" }
    }
}

function New-WordFixture {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string[]]$Paragraphs,
        [Parameter(Mandatory)][string]$FontName,
        [string]$FarEastFontName = '新細明體'
    )
    $word = $null
    $document = $null
    $selection = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $document = $word.Documents.Add()
        $normal = $document.Styles.Item(-1)
        $normal.Font.Name = $FontName
        $normal.Font.NameFarEast = $FarEastFontName
        $normal.Font.Size = 12
        Release-ComObject $normal

        $content = $Title + "`r`r" + ($Paragraphs -join "`r") + "`r"
        $contentRange = $document.Content
        $contentRange.Text = $content
        $contentRange.Font.Name = $FontName
        $contentRange.Font.NameFarEast = $FarEastFontName
        $contentRange.Font.Size = 12
        $titleRange = $document.Paragraphs.Item(1).Range
        $titleRange.Font.Size = 16
        $titleRange.Font.Bold = 1
        $titleRange.ParagraphFormat.Alignment = 1
        Release-ComObject $titleRange
        Release-ComObject $contentRange
        $document.SaveAs2([IO.Path]::GetFullPath($Path), 16)
        Write-Log "Created $Path"
    }
    finally {
        if ($null -ne $document) { try { $document.Close(0) } catch { Write-Log $_.Exception.Message } }
        if ($null -ne $word) { try { $word.Quit() } catch { Write-Log $_.Exception.Message } }
        foreach ($value in @($selection, $document, $word)) { Release-ComObject $value }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Add-SlideText {
    param(
        [Parameter(Mandatory)]$Slide,
        [Parameter(Mandatory)][string]$Text,
        [Parameter(Mandatory)][double]$Left,
        [Parameter(Mandatory)][double]$Top,
        [Parameter(Mandatory)][double]$Width,
        [Parameter(Mandatory)][double]$Height,
        [Parameter(Mandatory)][double]$FontSize,
        [int]$Alignment = 1,
        [switch]$Bold
    )
    $shape = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
    $shape.TextFrame.TextRange.Text = $Text
    $shape.TextFrame.TextRange.Font.Name = 'Arial'
    $shape.TextFrame.TextRange.Font.Size = $FontSize
    $shape.TextFrame.TextRange.Font.Bold = if ($Bold) { -1 } else { 0 }
    $shape.TextFrame.TextRange.ParagraphFormat.Alignment = $Alignment
    return $shape
}

function New-PowerPointFixture {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Lead,
        [Parameter(Mandatory)][string]$Marker,
        [Parameter(Mandatory)][string]$Definition
    )
    $powerPoint = $null
    $presentation = $null
    $slide = $null
    $shapes = [System.Collections.Generic.List[object]]::new()
    try {
        $powerPoint = New-Object -ComObject PowerPoint.Application
        $powerPoint.Visible = -1
        $presentation = $powerPoint.Presentations.Add(-1)
        $presentation.PageSetup.SlideWidth = 960
        $presentation.PageSetup.SlideHeight = 540
        $slide = $presentation.Slides.Add(1, 12)
        $shapes.Add((Add-SlideText -Slide $slide -Text $Title -Left 60 -Top 35 -Width 840 -Height 60 -FontSize 35 -Bold))
        $shapes.Add((Add-SlideText -Slide $slide -Text $Lead -Left 90 -Top 125 -Width 780 -Height 80 -FontSize 22 -Alignment 1))
        $shapes.Add((Add-SlideText -Slide $slide -Text $Marker -Left 180 -Top 230 -Width 600 -Height 55 -FontSize 28 -Alignment 2))
        $shapes.Add((Add-SlideText -Slide $slide -Text $Definition -Left 90 -Top 335 -Width 780 -Height 125 -FontSize 18 -Alignment 1))
        $presentation.SaveAs([IO.Path]::GetFullPath($Path), 24)
        Write-Log "Created $Path"
    }
    finally {
        if ($null -ne $presentation) { try { $presentation.Close() } catch { Write-Log $_.Exception.Message } }
        if ($null -ne $powerPoint) { try { $powerPoint.Quit() } catch { Write-Log $_.Exception.Message } }
        foreach ($shape in $shapes) { Release-ComObject $shape }
        foreach ($value in @($slide, $presentation, $powerPoint)) { Release-ComObject $value }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

try {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    if ($OnlyAutoClassification) {
        $fixtureBuilder = Join-Path $PSScriptRoot 'create_auto_classification_fixtures.py'
        $pythonHasDocx = $false
        try {
            & python -c 'import docx' 2>$null
            $pythonHasDocx = ($LASTEXITCODE -eq 0)
        }
        catch {
            Write-Log "Default Python cannot import python-docx: $($_.Exception.Message)"
        }

        if ($pythonHasDocx) {
            & python $fixtureBuilder --output-dir $OutputDirectory
        }
        elseif ($null -ne (Get-Command uv -ErrorAction SilentlyContinue)) {
            & uv run --with python-docx python $fixtureBuilder --output-dir $OutputDirectory
        }
        else {
            throw 'Generating automatic-classification fixtures requires python-docx. Install uv, or install the development dependency into the active Python environment.'
        }
        if ($LASTEXITCODE -ne 0) {
            throw "Automatic-classification fixture builder failed with exit code $LASTEXITCODE."
        }
        [ordered]@{ ok = $true; output_directory = [IO.Path]::GetFullPath($OutputDirectory); files = 2 } |
            ConvertTo-Json -Compress
        return
    }

    New-WordFixture -Path (Join-Path $OutputDirectory 'zh-auto-classification-draft.docx') `
        -Title '偏折角公式自動分類草稿' -FontName '新細明體' -Paragraphs @(
            '量測雜訊變異數 [[MATH id=measurement_variance tex=\sigma_v^2]] 用於描述重複量測的離散程度，其單位為 rad^2。',
            '扣除系統偏移量後的校正量可表示如下：',
            '[[MATH id=offset_correction tex=\widetilde{y}_i = y_i - b]]',
            '其中，y_i 為第 i 個取樣位置之原始量測值，單位為 rad；b 為系統偏移量，單位為 rad；\widetilde{y}_i 為校正後量測值，單位為 rad；i 為無因次取樣位置索引。',
            '單一取樣位置的偏折角增量可由下列模型求得：',
            '[[MATH id=local_deflection tex=\Delta \theta_{y,i} = K_i \Delta n_i]]',
            '其中，\Delta \theta_{y,i} 為第 i 個取樣位置在 y 方向之偏折角增量，單位為 rad；K_i 為第 i 個位置之靈敏度係數，單位為 rad；\Delta n_i 為無因次折射率差；i 為無因次取樣位置索引。',
            '由式 [[REF id=local_derivation target=local_deflection]] 得知，局部偏折角增量與折射率差成正比。',
            '為降低不同取樣數量造成的尺度差異，y 方向之累積偏折角定義如下：',
            '[[MATH id=cumulative_deflection tex=\theta_y = \frac{1}{N}\sum_{i=1}^{N}\Delta \theta_{y,i}]]',
            '其中，\theta_y 為 y 方向之累積偏折角，單位為 rad；\Delta \theta_{y,i} 為第 i 個取樣位置之偏折角增量，單位為 rad；i 為取樣位置索引；N 為取樣位置總數，i 與 N 皆為無因次正整數。',
            '其於 y 方向之累積偏折角可表示如式 [[REF id=cumulative_statement target=cumulative_deflection]] 所示。',
            '不同試次直接比較可能缺乏直觀的可比性。因此，如式 [[REF id=cumulative_comparison target=cumulative_deflection]] 所示，採用了依取樣位置總數正規化的累積偏折角。'
        )

    New-WordFixture -Path (Join-Path $OutputDirectory 'en-auto-classification-draft.docx') `
        -Title 'Automatic Equation Classification Draft' -FontName 'Times New Roman' -Paragraphs @(
            'The measurement-noise variance [[MATH id=measurement_variance tex=\sigma_v^2]] describes repeatability and has units of rad^2.',
            'The offset-corrected measurement is expressed as follows:',
            '[[MATH id=offset_correction tex=\widetilde{y}_i = y_i - b]]',
            'where y_i is the raw value at sample i in rad, b is the systematic offset in rad, \widetilde{y}_i is the corrected value in rad, and i is a dimensionless sample index.',
            'The local deflection-angle increment is obtained from the following model:',
            '[[MATH id=local_deflection tex=\Delta \theta_{y,i} = K_i \Delta n_i]]',
            'where \Delta \theta_{y,i} is the y-direction deflection-angle increment at sample i in rad, K_i is the sensitivity coefficient at sample i in rad, \Delta n_i is the dimensionless refractive-index difference, and i is a dimensionless sample index.',
            'Equation [[REF id=local_derivation target=local_deflection]] indicates that the local deflection increment is proportional to the refractive-index difference.',
            'To reduce scale differences caused by unequal sample counts, the cumulative y-direction deflection angle is defined as follows:',
            '[[MATH id=cumulative_deflection tex=\theta_y = \frac{1}{N}\sum_{i=1}^{N}\Delta \theta_{y,i}]]',
            'where \theta_y is the cumulative y-direction deflection angle in rad, \Delta \theta_{y,i} is the deflection-angle increment at sample i in rad, i is the sample index, and N is the sample count; i and N are dimensionless positive integers.',
            'Its cumulative deflection angle in the y direction can be expressed as shown in Eq. [[REF id=cumulative_statement target=cumulative_deflection]].',
            'Direct comparison may lack an intuitive basis when sample counts differ. Therefore, as shown in Eq. [[REF id=cumulative_comparison target=cumulative_deflection]], the sample-count-normalized cumulative deflection angle was adopted.'
        )

    New-WordFixture -Path (Join-Path $OutputDirectory 'zh-paper-draft.docx') `
        -Title '偏折角量測模型' -FontName '新細明體' -Paragraphs @(
            '影像量測所得的離散角度增量可累加為總偏折角，其定義如下：',
            '{{MATH:cumulative_deflection}}',
            '其中，θ_y 為 y 方向之累積偏折角，單位為 rad；Δθ_(y,i) 為第 i 個取樣位置在 y 方向之偏折角增量，單位為 rad；i 為取樣位置索引；N 為取樣位置總數，i 與 N 皆為無因次正整數。',
            '由式 {{EQREF:cumulative_deflection}} 得知，累積偏折角為各取樣位置偏折角增量之總和。',
            '不同試次的取樣位置總數可能不同，因此將累積偏折角正規化如下：',
            '{{MATH:normalized_deflection}}',
            '其中，η_y 為 y 方向之無因次正規化偏折角；θ_y 為 y 方向之累積偏折角，單位為 rad；N 為無因次取樣位置總數。',
            '不同試次直接比較可能缺乏直觀的可比性。因此，如式 {{EQREF:normalized_deflection}} 所示，採用了無因次正規化偏折角進行比較。'
        )

    New-WordFixture -Path (Join-Path $OutputDirectory 'en-paper-draft.docx') `
        -Title 'Discrete-Time Measurement Model' -FontName 'Times New Roman' -Paragraphs @(
            'The discrete-time state evolution is expressed as follows:',
            '{{MATH:state_transition}}',
            'where x_k is the state at sampling instant k, A is the state-transition coefficient, B is the input coefficient, u_k is the control input, and w_k is the process noise; k is a dimensionless integer time index.',
            'Equation {{EQREF:state_transition}} indicates that the next state depends on the current state, control input, and process noise.',
            'The corresponding observation model is expressed as follows:',
            '{{MATH:observation_model}}',
            'where y_k is the observation at sampling instant k, C is the observation coefficient, x_k is the state, and v_k is the measurement noise; k is a dimensionless integer time index.',
            'The measurement-noise variance {{MATH:measurement_variance}} has units equal to the squared observation unit.',
            'Direct comparison may lack an intuitive basis when observation scales differ. Therefore, as shown in Eq. {{EQREF:observation_model}}, the observation model was evaluated after unit-consistent normalization.'
        )

    New-PowerPointFixture -Path (Join-Path $OutputDirectory 'zh-presentation-draft.pptx') `
        -Title '顯熱傳遞模型' `
        -Lead '物體吸收之顯熱可表示如下：' `
        -Marker '{{MATH:sensible_heat}}' `
        -Definition '其中，Q 為顯熱量，單位為 J；m 為質量，單位為 kg；c_p 為定壓比熱容，單位為 J/(kg·K)；ΔT 為溫度變化量，單位為 K。'

    New-PowerPointFixture -Path (Join-Path $OutputDirectory 'en-presentation-draft.pptx') `
        -Title 'Prediction Error Metric' `
        -Lead 'The root-mean-square error is defined as follows:' `
        -Marker '{{MATH:root_mean_square_error}}' `
        -Definition 'where RMSE is the root-mean-square error in the observation unit, N is the dimensionless sample count, i is the sample index, y_i is the observed value, and ŷ_i is the predicted value.'

    [ordered]@{ ok = $true; output_directory = [IO.Path]::GetFullPath($OutputDirectory); files = 6 } |
        ConvertTo-Json -Compress
}
catch {
    Write-Log $_.Exception.ToString()
    throw
}
