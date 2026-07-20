[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet('probe', 'probe-pptx', 'configure-defaults', 'render', 'validate', 'update', 'render-pptx', 'validate-pptx')]
    [string]$Action,

    [string]$InputPath,
    [string]$OutputPath,
    [string]$ManifestPath,
    [switch]$Overwrite
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:PluginRoot = Split-Path -Parent $PSScriptRoot
$script:PackagedDefaultsPath = Join-Path $script:PluginRoot 'config\defaults.json'
$script:UserConfigDirectory = Join-Path $env:APPDATA 'MathTypeForWordAgent'
$script:UserDefaultsPath = Join-Path $script:UserConfigDirectory 'defaults.json'
$script:WordCommandsKey = 'HKCU:\Software\Design Science\DSMT7\WordCommands'
$script:Word = $null
$script:Document = $null
$script:Selection = $null
$script:PowerPoint = $null
$script:Presentation = $null
$script:HeldComObjects = [System.Collections.Generic.List[object]]::new()
$script:OriginalNumberWarning = $null
$script:OriginalReferenceWarning = $null
$script:AutomationPollMilliseconds = 250
$script:AutomationTimeoutMilliseconds = if ($env:MATHTYPE_AUTOMATION_TIMEOUT_MS) {
    [int]$env:MATHTYPE_AUTOMATION_TIMEOUT_MS
}
else { 15000 }

if (-not ('MathTypeOleData' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Runtime.InteropServices;
public static class MathTypeOleData {
    private const uint GMEM_MOVEABLE_ZEROINIT = 0x0042;

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern uint RegisterClipboardFormat(string format);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GlobalAlloc(uint flags, UIntPtr bytes);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GlobalLock(IntPtr memory);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GlobalUnlock(IntPtr memory);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GlobalFree(IntPtr memory);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern UIntPtr GlobalSize(IntPtr memory);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool OpenClipboard(IntPtr owner);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool EmptyClipboard();

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetClipboardData(uint format, IntPtr memory);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool CloseClipboard();

    [DllImport("ole32.dll")]
    private static extern void ReleaseStgMedium(ref STGMEDIUM medium);

    private static string WrapMathML(string mathML) {
        if (string.IsNullOrWhiteSpace(mathML)) throw new ArgumentException("MathML is empty.", "mathML");
        return mathML.TrimStart().StartsWith("<?xml", StringComparison.Ordinal)
            ? mathML
            : "<?xml version='1.0'?><!-- MathType@Translator@5@5@MathML2 (Clipboard).tdl@MathML 2.0 (Clipboard)@ --><html>" + mathML + "</html><!-- MathType@End@5@5@ -->";
    }

    private static void PutClipboardBytes(uint format, byte[] bytes) {
        IntPtr memory = GlobalAlloc(GMEM_MOVEABLE_ZEROINIT, (UIntPtr)bytes.Length);
        if (memory == IntPtr.Zero) throw new OutOfMemoryException("GlobalAlloc failed for MathML data.");
        try {
            IntPtr target = GlobalLock(memory);
            if (target == IntPtr.Zero) throw new InvalidOperationException("GlobalLock failed for MathML data.");
            try { Marshal.Copy(bytes, 0, target, bytes.Length); }
            finally { GlobalUnlock(memory); }
            if (SetClipboardData(format, memory) == IntPtr.Zero)
                throw new InvalidOperationException("SetClipboardData failed for MathML data.");
            memory = IntPtr.Zero;
        }
        finally { if (memory != IntPtr.Zero) GlobalFree(memory); }
    }

    public static void SetMathMLClipboard(string mathML) {
        byte[] bytes = Encoding.UTF8.GetBytes(WrapMathML(mathML) + "\0");
        byte[] unicodeText = Encoding.Unicode.GetBytes(mathML + "\0");
        bool opened = false;
        for (int attempt = 0; attempt < 20 && !opened; attempt++) {
            opened = OpenClipboard(IntPtr.Zero);
            if (!opened) System.Threading.Thread.Sleep(50);
        }
        if (!opened) throw new InvalidOperationException("OpenClipboard failed for MathML data.");
        try {
            if (!EmptyClipboard()) throw new InvalidOperationException("EmptyClipboard failed for MathML data.");
            PutClipboardBytes(RegisterClipboardFormat("MathML Presentation"), bytes);
            PutClipboardBytes(RegisterClipboardFormat("MathML"), bytes);
            PutClipboardBytes(RegisterClipboardFormat("application/mathml+xml"), bytes);
            PutClipboardBytes(13, unicodeText); // CF_UNICODETEXT: MathType 7 parses bare MathML on paste.
        }
        finally { CloseClipboard(); }
    }

    public static string GetMathML(object oleObject) {
        if (oleObject == null) throw new ArgumentNullException("oleObject");
        IDataObject dataObject = oleObject as IDataObject;
        if (dataObject == null) throw new InvalidOperationException("MathType OLE object does not expose IDataObject.");
        string[] formatNames = { "MathML Presentation", "MathML", "application/mathml+xml" };
        foreach (string formatName in formatNames) {
            FORMATETC format = new FORMATETC {
                cfFormat = unchecked((short)RegisterClipboardFormat(formatName)),
                dwAspect = DVASPECT.DVASPECT_CONTENT,
                lindex = -1,
                ptd = IntPtr.Zero,
                tymed = TYMED.TYMED_HGLOBAL
            };
            int query;
            try { query = dataObject.QueryGetData(ref format); }
            catch (COMException ex) { query = ex.HResult; }
            if (query != 0) continue;
            STGMEDIUM medium;
            dataObject.GetData(ref format, out medium);
            try {
                IntPtr source = GlobalLock(medium.unionmember);
                if (source == IntPtr.Zero) throw new InvalidOperationException("GlobalLock failed while reading MathML.");
                try {
                    ulong byteCount64 = GlobalSize(medium.unionmember).ToUInt64();
                    int byteCount = byteCount64 > Int32.MaxValue ? Int32.MaxValue : (int)byteCount64;
                    byte[] bytes = new byte[byteCount];
                    Marshal.Copy(source, bytes, 0, byteCount);
                    int length = Array.IndexOf(bytes, (byte)0);
                    if (length < 0) length = bytes.Length;
                    return Encoding.UTF8.GetString(bytes, 0, length);
                }
                finally { GlobalUnlock(medium.unionmember); }
            }
            finally { ReleaseStgMedium(ref medium); }
        }
        throw new InvalidOperationException("MathType OLE object exposes no readable MathML format.");
    }
}
'@
}

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level,
        [Parameter(Mandatory)]
        [string]$Message
    )
    [Console]::Error.WriteLine("[$Level] $Message")
}

function Write-JsonResult {
    param([Parameter(Mandatory)]$Value)
    $Value | ConvertTo-Json -Depth 20 -Compress
}

function Resolve-ExistingFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Label
    )
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "$Label does not exist: $Path"
    }
    return (Resolve-Path -LiteralPath $Path).Path
}

function Resolve-OutputFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Extension = '.docx'
    )
    $fullPath = [IO.Path]::GetFullPath($Path)
    if ([IO.Path]::GetExtension($fullPath) -ine $Extension) {
        throw "OutputPath must use the $Extension extension: $fullPath"
    }
    $parent = Split-Path -Parent $fullPath
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    if ((Test-Path -LiteralPath $fullPath) -and -not $Overwrite) {
        throw "OutputPath already exists. Pass -Overwrite to replace it: $fullPath"
    }
    return $fullPath
}

function Release-ComObject {
    param($Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) } catch {}
    }
}

function Start-Word {
    param([switch]$ReadOnly, [string]$DocumentPath)
    $before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $script:Word = New-Object -ComObject Word.Application
    $script:Word.Visible = $false
    $script:Word.DisplayAlerts = 0
    if ([string]::IsNullOrWhiteSpace($DocumentPath)) {
        $script:Document = $script:Word.Documents.Add()
    }
    else {
        $script:Document = $script:Word.Documents.Open($DocumentPath, $false, [bool]$ReadOnly)
    }
    $newProcesses = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id })
    if ($newProcesses.Count -eq 1) { Write-Log -Level INFO -Message "WORD_PID=$($newProcesses[0].Id)" }
    elseif ($newProcesses.Count -gt 1) { Write-Log -Level WARN -Message "Multiple new Word PIDs observed: $($newProcesses.Id -join ', ')" }
    $script:Selection = $script:Word.Selection
}

function Start-PowerPoint {
    param([Parameter(Mandatory)][string]$PresentationPath, [switch]$ReadOnly)
    $before = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $script:PowerPoint = New-Object -ComObject PowerPoint.Application
    $script:PowerPoint.Visible = -1
    $script:Presentation = $script:PowerPoint.Presentations.Open(
        $PresentationPath,
        [bool]$ReadOnly,
        $false,
        (-not [bool]$ReadOnly)
    )
    $newProcesses = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id })
    if ($newProcesses.Count -eq 1) {
        Write-Log -Level INFO -Message "POWERPOINT_PID=$($newProcesses[0].Id)"
    }
}

function Stop-Word {
    param([switch]$Save)
    if ($null -ne $script:Document) {
        try { $script:Document.Close($(if ($Save) { -1 } else { 0 })) } catch {
            Write-Log -Level WARN -Message "Could not close Word document cleanly: $($_.Exception.Message)"
        }
    }
    if ($null -ne $script:Word) {
        try { $script:Word.Quit() } catch {
            Write-Log -Level WARN -Message "Could not quit Word cleanly: $($_.Exception.Message)"
        }
    }
    foreach ($value in $script:HeldComObjects) { Release-ComObject $value }
    Release-ComObject $script:Selection
    Release-ComObject $script:Document
    Release-ComObject $script:Word
    $script:HeldComObjects.Clear()
    $script:Selection = $null
    $script:Document = $null
    $script:Word = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Stop-PowerPoint {
    if ($null -ne $script:Presentation) {
        try { $script:Presentation.Close() } catch {
            Write-Log -Level WARN -Message "Could not close PowerPoint presentation cleanly: $($_.Exception.Message)"
        }
    }
    if ($null -ne $script:PowerPoint) {
        try { $script:PowerPoint.Quit() } catch {
            Write-Log -Level WARN -Message "Could not quit PowerPoint cleanly: $($_.Exception.Message)"
        }
    }
    Release-ComObject $script:Presentation
    Release-ComObject $script:PowerPoint
    $script:Presentation = $null
    $script:PowerPoint = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Set-TemporaryMathTypePreferences {
    if (-not (Test-Path -LiteralPath $script:WordCommandsKey)) {
        throw 'MathType 7 WordCommands registry key was not found.'
    }
    $properties = Get-ItemProperty -LiteralPath $script:WordCommandsKey
    $script:OriginalNumberWarning = $properties.NoEqnNumWarningDlg
    $script:OriginalReferenceWarning = $properties.NoInsertEqnRefDlg
    Set-ItemProperty -LiteralPath $script:WordCommandsKey -Name NoEqnNumWarningDlg -Value 1
    Set-ItemProperty -LiteralPath $script:WordCommandsKey -Name NoInsertEqnRefDlg -Value 1
}

function Restore-MathTypePreferences {
    if ($null -ne $script:OriginalNumberWarning -and (Test-Path -LiteralPath $script:WordCommandsKey)) {
        Set-ItemProperty -LiteralPath $script:WordCommandsKey -Name NoEqnNumWarningDlg -Value $script:OriginalNumberWarning
    }
    if ($null -ne $script:OriginalReferenceWarning -and (Test-Path -LiteralPath $script:WordCommandsKey)) {
        Set-ItemProperty -LiteralPath $script:WordCommandsKey -Name NoInsertEqnRefDlg -Value $script:OriginalReferenceWarning
    }
}

function Read-Manifest {
    param([Parameter(Mandatory)][string]$Path)
    $resolved = Resolve-ExistingFile -Path $Path -Label 'ManifestPath'
    $manifest = Get-Content -Raw -LiteralPath $resolved | ConvertFrom-Json
    if ($manifest.schema_version -ne 1) { throw 'Manifest schema_version must be 1.' }
    if ($null -eq $manifest.equations) { throw 'Manifest must contain an equations array.' }
    if ($null -eq $manifest.references) {
        $manifest | Add-Member -NotePropertyName references -NotePropertyValue @()
    }
    return $manifest
}

function Assert-Manifest {
    param([Parameter(Mandatory)]$Manifest)
    $ids = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $markers = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($equation in @($Manifest.equations)) {
        if ([string]::IsNullOrWhiteSpace($equation.id)) { throw 'Every equation requires a non-empty id.' }
        if (-not $ids.Add([string]$equation.id)) { throw "Duplicate equation id: $($equation.id)" }
        if ([string]::IsNullOrWhiteSpace($equation.marker)) { throw "Equation '$($equation.id)' requires a marker." }
        if (-not $markers.Add([string]$equation.marker)) { throw "Duplicate marker: $($equation.marker)" }
        if ([string]::IsNullOrWhiteSpace($equation.tex)) { throw "Equation '$($equation.id)' requires TeX." }
        if (@('inline', 'display') -notcontains [string]$equation.layout) {
            throw "Equation '$($equation.id)' layout must be inline or display."
        }
        if ([bool]$equation.numbered -and [string]$equation.layout -ne 'display') {
            throw "Equation '$($equation.id)' cannot be numbered unless layout is display."
        }
        if ([string]$equation.tex -match '[\r\n]') { throw "Equation '$($equation.id)' TeX must be one line." }
    }
    foreach ($reference in @($Manifest.references)) {
        if ([string]::IsNullOrWhiteSpace($reference.marker)) { throw 'Every reference requires a marker.' }
        if (-not $markers.Add([string]$reference.marker)) { throw "Duplicate marker: $($reference.marker)" }
        if (-not $ids.Contains([string]$reference.target)) {
            throw "Reference target does not exist: $($reference.target)"
        }
        $target = @($Manifest.equations | Where-Object id -eq $reference.target)[0]
        if (-not [bool]$target.numbered) { throw "Reference target is not numbered: $($reference.target)" }
    }
}

function Read-PresentationManifest {
    param([Parameter(Mandatory)][string]$Path)
    $resolved = Resolve-ExistingFile -Path $Path -Label 'ManifestPath'
    $manifest = Get-Content -Raw -LiteralPath $resolved | ConvertFrom-Json
    if ($manifest.schema_version -ne 1) { throw 'Presentation manifest schema_version must be 1.' }
    if ($null -eq $manifest.equations) { throw 'Presentation manifest must contain an equations array.' }
    return $manifest
}

function Assert-PresentationManifest {
    param([Parameter(Mandatory)]$Manifest)
    $ids = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $markers = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($equation in @($Manifest.equations)) {
        if ([string]::IsNullOrWhiteSpace($equation.id)) { throw 'Every presentation equation requires a non-empty id.' }
        if ([string]$equation.id -notmatch '^[A-Za-z][A-Za-z0-9_-]*$') {
            throw "Presentation equation id must be stable ASCII: $($equation.id)"
        }
        if (-not $ids.Add([string]$equation.id)) { throw "Duplicate presentation equation id: $($equation.id)" }
        if ([string]::IsNullOrWhiteSpace($equation.marker)) { throw "Presentation equation '$($equation.id)' requires a marker." }
        if (-not $markers.Add([string]$equation.marker)) { throw "Duplicate presentation marker: $($equation.marker)" }
        if ([string]::IsNullOrWhiteSpace($equation.mathml)) { throw "Presentation equation '$($equation.id)' requires Presentation MathML." }
        if ([string]$equation.mathml -notmatch '^\s*<math\b' -or [string]$equation.mathml -notmatch '</math>\s*$') {
            throw "Presentation equation '$($equation.id)' mathml must contain one complete MathML math element."
        }
        if ($null -ne $equation.height_points) {
            $height = [double]$equation.height_points
            if ($height -lt 12 -or $height -gt 200) {
                throw "Presentation equation '$($equation.id)' height_points must be between 12 and 200."
            }
        }
    }
}

function Find-UniqueMarkerRange {
    param([Parameter(Mandatory)][string]$Marker)
    $search = $script:Document.Content.Duplicate
    $matches = @()
    try {
        while ($search.Start -lt $search.End) {
            $find = $search.Find
            $find.ClearFormatting()
            $find.Text = $Marker
            $find.Forward = $true
            $find.Wrap = 0
            $find.MatchWildcards = $false
            if (-not $find.Execute()) { break }
            $matches += [ordered]@{ start = $search.Start; end = $search.End }
            $next = $search.End
            $search.SetRange($next, $script:Document.Content.End)
            Release-ComObject $find
        }
    }
    finally { Release-ComObject $search }
    if ($matches.Count -ne 1) {
        throw "Marker must occur exactly once; found $($matches.Count): $Marker"
    }
    return $script:Document.Range($matches[0].start, $matches[0].end)
}

function Get-MathTypeShapeNear {
    param([Parameter(Mandatory)][int]$Position)
    $best = $null
    $bestDistance = [int]::MaxValue
    for ($index = 1; $index -le $script:Document.InlineShapes.Count; $index++) {
        $shape = $script:Document.InlineShapes.Item($index)
        $progId = $null
        try { $progId = $shape.OLEFormat.ProgID } catch {}
        if ($progId -eq 'Equation.DSMT4') {
            $distance = [Math]::Abs($shape.Range.Start - $Position)
            if ($distance -lt $bestDistance) {
                Release-ComObject $best
                $best = $shape
                $bestDistance = $distance
                continue
            }
        }
        Release-ComObject $shape
    }
    if ($null -eq $best -or $bestDistance -gt 5) {
        Release-ComObject $best
        throw "MathType Toggle TeX did not create an Equation.DSMT4 object near position $Position."
    }
    return $best
}

function Find-NearestNumberField {
    param([Parameter(Mandatory)][int]$Position)
    $deadline = [DateTime]::UtcNow.AddMilliseconds($script:AutomationTimeoutMilliseconds)
    do {
        $best = $null
        $bestDistance = [int]::MaxValue
        for ($index = 1; $index -le $script:Document.Fields.Count; $index++) {
            $field = $script:Document.Fields.Item($index)
            if ($field.Code.Text -match '^\s*MACROBUTTON\s+MTPlaceRef\b') {
                $distance = [Math]::Abs($field.Code.Start - $Position)
                if ($distance -lt $bestDistance) {
                    Release-ComObject $best
                    $best = $field
                    $bestDistance = $distance
                    continue
                }
            }
            Release-ComObject $field
        }
        if ($null -ne $best) { return $best }
        Start-Sleep -Milliseconds $script:AutomationPollMilliseconds
    } while ([DateTime]::UtcNow -lt $deadline)
    throw "MathType did not create a native MTPlaceRef number field within $($script:AutomationTimeoutMilliseconds) ms."
}

function Normalize-SimpleNumberField {
    param([Parameter(Mandatory)]$NumberField)
    $nestedToDelete = @()
    for ($index = 1; $index -le $script:Document.Fields.Count; $index++) {
        $field = $script:Document.Fields.Item($index)
        $code = $field.Code.Text
        if ($code -match '^\s*SEQ\s+(MTSec|MTChap)\s+\\c\b') {
            if ($field.Code.Start -ge $NumberField.Code.Start -and $field.Code.End -le $NumberField.Code.End) {
                $nestedToDelete += $field
                continue
            }
        }
        Release-ComObject $field
    }
    foreach ($field in $nestedToDelete) {
        $field.Delete()
        Release-ComObject $field
    }

    $equationCurrent = $null
    for ($index = 1; $index -le $script:Document.Fields.Count; $index++) {
        $field = $script:Document.Fields.Item($index)
        if (
            $field.Code.Text -match '^\s*SEQ\s+MTEqn\s+\\c\b' -and
            $field.Code.Start -ge $NumberField.Code.Start -and
            $field.Code.End -le $NumberField.Code.End
        ) {
            $equationCurrent = $field
            break
        }
        Release-ComObject $field
    }
    if ($null -eq $equationCurrent) { throw 'Native number field is missing SEQ MTEqn \\c.' }

    $leftParenthesis = $NumberField.Code.Duplicate
    $leftFind = $leftParenthesis.Find
    $leftFind.ClearFormatting()
    $leftFind.Text = '('
    if (-not $leftFind.Execute()) { throw 'Native MathType number is not enclosed in parentheses.' }
    $cleanupEnd = $equationCurrent.Code.Start - 1
    if ($cleanupEnd -gt $leftParenthesis.End) {
        $separatorRange = $script:Document.Range($leftParenthesis.End, $cleanupEnd)
        $separatorRange.Delete() | Out-Null
        Release-ComObject $separatorRange
    }
    $dotRange = $NumberField.Code.Duplicate
    $dotFind = $dotRange.Find
    $dotFind.ClearFormatting()
    $dotFind.Text = '.'
    while ($dotFind.Execute()) {
        $dotRange.Delete() | Out-Null
        Release-ComObject $dotFind
        Release-ComObject $dotRange
        $dotRange = $NumberField.Code.Duplicate
        $dotFind = $dotRange.Find
        $dotFind.ClearFormatting()
        $dotFind.Text = '.'
    }
    Release-ComObject $dotFind
    Release-ComObject $dotRange
    Release-ComObject $leftFind
    Release-ComObject $leftParenthesis
    Release-ComObject $equationCurrent
    $NumberField.Update() | Out-Null
}

function Convert-Equation {
    param([Parameter(Mandatory)]$Equation)
    Write-Log -Level INFO -Message "Converting equation '$($Equation.id)' as $($Equation.layout), numbered=$([bool]$Equation.numbered)."
    $markerRange = Find-UniqueMarkerRange -Marker ([string]$Equation.marker)
    $position = $markerRange.Start
    $delimited = if ([string]$Equation.layout -eq 'inline') {
        '$' + [string]$Equation.tex + '$'
    }
    else {
        '\[' + [string]$Equation.tex + '\]'
    }
    $markerRange.Text = $delimited
    $markerRange.SetRange($position, $position + $delimited.Length)
    $markerRange.Select()
    $before = $script:Document.InlineShapes.Count
    Write-Log -Level INFO -Message "Calling MTCommand_TeXToggle for '$($Equation.id)'."
    $null = $script:Word.Run('MTCommand_TeXToggle')
    Write-Log -Level INFO -Message "Toggle TeX returned for '$($Equation.id)'."
    if ($script:Document.InlineShapes.Count -ne ($before + 1)) {
        throw "Toggle TeX did not add exactly one MathType object for '$($Equation.id)'."
    }
    $shape = Get-MathTypeShapeNear -Position $position
    Release-ComObject $markerRange

    $numberField = $null
    if ([bool]$Equation.numbered) {
        $paragraph = $shape.Range.Paragraphs.Item(1)
        $insertPosition = $paragraph.Range.End - 1
        $script:Selection.SetRange($insertPosition, $insertPosition)
        $script:Selection.TypeText("`t")
        $insertPosition = $script:Selection.End
        Write-Log -Level INFO -Message "Calling MTCommand_InsertEqnNum for '$($Equation.id)'."
        $null = $script:Word.Run('MTCommand_InsertEqnNum')
        Write-Log -Level INFO -Message "Equation-number insertion returned for '$($Equation.id)'."
        $numberField = Find-NearestNumberField -Position $insertPosition
        Normalize-SimpleNumberField -NumberField $numberField
        Write-Log -Level INFO -Message "Normalized native number for '$($Equation.id)' to simple parenthesized format."
        Release-ComObject $paragraph
    }
    Release-ComObject $shape
    return $numberField
}

function Insert-NativeReference {
    param(
        [Parameter(Mandatory)]$Reference,
        [Parameter(Mandatory)]$TargetField
    )
    $markerRange = Find-UniqueMarkerRange -Marker ([string]$Reference.marker)
    Write-Log -Level INFO -Message "Inserting native reference '$($Reference.marker)' -> '$($Reference.target)'."
    $position = $markerRange.Start
    $markerRange.Text = ''
    $script:Selection.SetRange($position, $position)
    $null = $script:Word.Run('MTCommand_InsertEqnRef')
    Write-Log -Level INFO -Message "Reference placeholder command returned for '$($Reference.marker)'."
    if (-not $script:Document.Bookmarks.Exists('MTReference')) {
        throw "MathType did not create the MTReference placeholder for '$($Reference.marker)'."
    }
    $TargetField.DoClick()
    Write-Log -Level INFO -Message "MTPlaceRef returned for '$($Reference.marker)'."
    if ($script:Document.Bookmarks.Exists('MTReference')) {
        throw "MathType did not resolve the reference marker '$($Reference.marker)'."
    }
    Release-ComObject $markerRange
}

function Save-DocumentAtomically {
    param([Parameter(Mandatory)][string]$Destination)
    $temporary = Join-Path (Split-Path -Parent $Destination) ('.' + [IO.Path]::GetFileNameWithoutExtension($Destination) + '.' + [Guid]::NewGuid().ToString('N') + '.tmp.docx')
    try {
        Write-Log -Level INFO -Message "Saving temporary DOCX: $temporary"
        $script:Document.SaveAs2($temporary, 16)
        Write-Log -Level INFO -Message 'Word SaveAs2 returned.'
        Stop-Word
        Move-Item -LiteralPath $temporary -Destination $Destination -Force
        Write-Log -Level INFO -Message "Moved completed DOCX to: $Destination"
    }
    finally {
        if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force }
    }
}

function Find-PresentationMarkerShape {
    param([Parameter(Mandatory)][string]$Marker)
    $matches = @()
    for ($slideIndex = 1; $slideIndex -le $script:Presentation.Slides.Count; $slideIndex++) {
        $slide = $script:Presentation.Slides.Item($slideIndex)
        for ($shapeIndex = 1; $shapeIndex -le $slide.Shapes.Count; $shapeIndex++) {
            $shape = $slide.Shapes.Item($shapeIndex)
            $text = $null
            try {
                if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                    $text = [string]$shape.TextFrame.TextRange.Text
                }
            }
            catch {
                Write-Log -Level WARN -Message "Could not inspect text in slide $slideIndex shape ${shapeIndex}: $($_.Exception.Message)"
            }
            if ($null -ne $text -and $text.Trim() -eq $Marker) {
                $matches += [ordered]@{ slide = $slide; shape = $shape }
                continue
            }
            Release-ComObject $shape
        }
        if (-not ($matches | Where-Object { $_.slide -eq $slide })) { Release-ComObject $slide }
    }
    if ($matches.Count -ne 1) {
        foreach ($match in $matches) {
            Release-ComObject $match.shape
            Release-ComObject $match.slide
        }
        throw "Presentation marker must be the complete text of exactly one top-level shape; found $($matches.Count): $Marker"
    }
    return $matches[0]
}

function Get-MathMLTextSignature {
    param([Parameter(Mandatory)][string]$MathML)
    try { [xml]$document = $MathML }
    catch { throw "Invalid MathML XML: $($_.Exception.Message)" }
    $text = [string]$document.DocumentElement.InnerText
    $normalized = [regex]::Replace($text, '\s+', '')
    # MathType may encode scalable delimiters as template structure instead of
    # literal MathML text nodes when it round-trips the equation.
    return [regex]::Replace($normalized, '[\(\)\[\]\{\}]', '')
}

function Add-MathTypeEquationToSlide {
    param(
        [Parameter(Mandatory)]$Equation,
        [Parameter(Mandatory)]$Slide,
        [Parameter(Mandatory)][double]$Top
    )
    $height = if ($null -ne $Equation.height_points) { [double]$Equation.height_points } else { 32.0 }
    $shape = $null
    $oleObject = $null
    $shell = $null
    try {
        $shape = $Slide.Shapes.AddOLEObject(
            0,
            $Top,
            -1,
            -1,
            'Equation.DSMT4',
            '',
            $false,
            '',
            0,
            '',
            $false
        )
        $expectedSignature = Get-MathMLTextSignature -MathML ([string]$Equation.mathml)
        $converted = $false
        $lastFailure = 'No conversion attempt ran.'
        for ($attempt = 1; $attempt -le 3 -and -not $converted; $attempt++) {
            try {
                [MathTypeOleData]::SetMathMLClipboard([string]$Equation.mathml)
                $shape.OLEFormat.Activate()
                $deadline = [DateTime]::UtcNow.AddMilliseconds($script:AutomationTimeoutMilliseconds)
                $mathTypeProcess = $null
                while ([DateTime]::UtcNow -lt $deadline) {
                    $mathTypeProcess = Get-Process MathType -ErrorAction SilentlyContinue |
                        Where-Object { $_.MainWindowHandle -ne 0 } |
                        Select-Object -First 1
                    if ($null -ne $mathTypeProcess) { break }
                    Start-Sleep -Milliseconds $script:AutomationPollMilliseconds
                }
                if ($null -eq $mathTypeProcess) { throw 'MathType editor window did not appear.' }
                $shell = New-Object -ComObject WScript.Shell
                if (-not $shell.AppActivate([int]$mathTypeProcess.Id)) {
                    throw "Could not activate MathType editor process $($mathTypeProcess.Id)."
                }
                Start-Sleep -Milliseconds 700
                $shell.SendKeys('^a')
                Start-Sleep -Milliseconds 200
                $shell.SendKeys('^v')
                Start-Sleep -Milliseconds 1200
                $shell.SendKeys('%{F4}')
                Start-Sleep -Milliseconds 2500

                # RunForConversion refreshes PowerPoint's cached OLE preview after the
                # embedded MathType editor has committed the MathML payload.
                $shape.OLEFormat.DoVerb(2)
                Start-Sleep -Milliseconds 500
                $oleObject = $shape.OLEFormat.Object
                $embeddedMathML = [MathTypeOleData]::GetMathML($oleObject)
                if ($embeddedMathML -notmatch '<(?:[A-Za-z_][\w.-]*:)?math\b') {
                    throw 'Embedded MathType object did not expose a MathML root after editing.'
                }
                if ($embeddedMathML.Contains([char]0xFFFD) -or $embeddedMathML -match '&#x0*FFFD;') {
                    throw 'Embedded MathType object contains a Unicode replacement character.'
                }
                $actualSignature = Get-MathMLTextSignature -MathML $embeddedMathML
                if ([string]::IsNullOrWhiteSpace($actualSignature) -or $actualSignature -ne $expectedSignature) {
                    throw "MathML text signature mismatch: expected '$expectedSignature', got '$actualSignature'."
                }
                $converted = $true
            }
            catch {
                $lastFailure = $_.Exception.Message
                Write-Log -Level WARN -Message "PowerPoint MathType conversion attempt $attempt/3 failed for '$($Equation.id)': $lastFailure"
            }
            finally {
                Release-ComObject $shell
                Release-ComObject $oleObject
                $shell = $null
                $oleObject = $null
            }
        }
        if (-not $converted) {
            throw "MathType editor failed to commit the requested MathML after 3 attempts: $lastFailure"
        }
        $shape.LockAspectRatio = -1
        $shape.Height = $height
        $shape.Left = ($script:Presentation.PageSetup.SlideWidth - $shape.Width) / 2
        $shape.Top = $Top
        $shape.Name = "MathType_$($Equation.id)"
        $shape.AlternativeText = "MathType equation: $($Equation.id)"
        return $shape
    }
    catch {
        if ($null -ne $shape) { try { $shape.Delete() } catch {} }
        Release-ComObject $shape
        throw "Direct PowerPoint MathType OLE insertion failed for '$($Equation.id)': $($_.Exception.Message)"
    }
    finally {
        Release-ComObject $shell
        Release-ComObject $oleObject
    }
}

function Save-PresentationAtomically {
    param([Parameter(Mandatory)][string]$Destination)
    $temporary = Join-Path (Split-Path -Parent $Destination) ('.' + [IO.Path]::GetFileNameWithoutExtension($Destination) + '.' + [Guid]::NewGuid().ToString('N') + '.tmp.pptx')
    try {
        Write-Log -Level INFO -Message "Saving temporary PPTX: $temporary"
        $script:Presentation.SaveAs($temporary, 24)
        Stop-PowerPoint
        Move-Item -LiteralPath $temporary -Destination $Destination -Force
    }
    finally {
        if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force }
    }
}

function Invoke-RenderPptx {
    if ([string]::IsNullOrWhiteSpace($InputPath)) { throw 'render-pptx requires InputPath.' }
    if ([string]::IsNullOrWhiteSpace($OutputPath)) { throw 'render-pptx requires OutputPath.' }
    if ([string]::IsNullOrWhiteSpace($ManifestPath)) { throw 'render-pptx requires ManifestPath.' }
    $input = Resolve-ExistingFile -Path $InputPath -Label 'InputPath'
    if ([IO.Path]::GetExtension($input) -ine '.pptx') { throw 'InputPath must be a .pptx file.' }
    $output = Resolve-OutputFile -Path $OutputPath -Extension '.pptx'
    $manifest = Read-PresentationManifest -Path $ManifestPath
    Assert-PresentationManifest -Manifest $manifest

    try {
        Start-PowerPoint -PresentationPath $input
        foreach ($equation in @($manifest.equations)) {
            $marker = Find-PresentationMarkerShape -Marker ([string]$equation.marker)
            $top = [double]$marker.shape.Top
            $pastedShape = Add-MathTypeEquationToSlide -Equation $equation -Slide $marker.slide -Top $top
            $marker.shape.Delete()
            Release-ComObject $pastedShape
            Release-ComObject $marker.shape
            Release-ComObject $marker.slide
        }
        Save-PresentationAtomically -Destination $output
    }
    finally {
        Stop-PowerPoint
    }
    Write-JsonResult ([ordered]@{
        ok = $true
        action = 'render-pptx'
        input_path = $input
        output_path = $output
        equations = @($manifest.equations).Count
        object_type = 'Equation.DSMT4 floating OLE'
        numbering_and_references = 'Not available through the MathType 7 PowerPoint integration.'
    })
}

function Get-PresentationValidationReport {
    param([Parameter(Mandatory)][string]$PresentationPath, [Parameter(Mandatory)]$Manifest)
    $errors = [System.Collections.Generic.List[string]]::new()
    $mathTypeObjects = 0
    $verifiedMathML = 0
    $namedObjects = @()
    $unresolvedMarkers = @()
    $slideWidth = [double]$script:Presentation.PageSetup.SlideWidth
    $expectedSignatures = @{}
    foreach ($equation in @($Manifest.equations)) {
        $expectedSignatures["MathType_$($equation.id)"] = Get-MathMLTextSignature -MathML ([string]$equation.mathml)
    }
    for ($slideIndex = 1; $slideIndex -le $script:Presentation.Slides.Count; $slideIndex++) {
        $slide = $script:Presentation.Slides.Item($slideIndex)
        for ($shapeIndex = 1; $shapeIndex -le $slide.Shapes.Count; $shapeIndex++) {
            $shape = $slide.Shapes.Item($shapeIndex)
            $progId = $null
            try { $progId = $shape.OLEFormat.ProgID } catch {}
            if ($progId -eq 'Equation.DSMT4') {
                $mathTypeObjects++
                $namedObjects += [string]$shape.Name
                $centerError = [Math]::Abs(($shape.Left + ($shape.Width / 2)) - ($slideWidth / 2))
                if ($centerError -gt 1.0) { $errors.Add("$($shape.Name) on slide $slideIndex is not horizontally centered.") }
                if ($expectedSignatures.ContainsKey([string]$shape.Name)) {
                    $oleObject = $null
                    try {
                        $oleObject = $shape.OLEFormat.Object
                        $embeddedMathML = [MathTypeOleData]::GetMathML($oleObject)
                        $actualSignature = Get-MathMLTextSignature -MathML $embeddedMathML
                        if ($embeddedMathML.Contains([char]0xFFFD) -or $embeddedMathML -match '&#x0*FFFD;') {
                            $errors.Add("$($shape.Name) contains a Unicode replacement character.")
                        }
                        elseif ($actualSignature -ne $expectedSignatures[[string]$shape.Name]) {
                            $errors.Add("$($shape.Name) MathML content does not match its manifest equation.")
                        }
                        else { $verifiedMathML++ }
                    }
                    catch { $errors.Add("Could not verify MathML in $($shape.Name): $($_.Exception.Message)") }
                    finally { Release-ComObject $oleObject }
                }
            }
            try {
                if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                    $text = [string]$shape.TextFrame.TextRange.Text
                    if ($text -match '\{\{MATH:') { $unresolvedMarkers += $text.Trim() }
                }
            }
            catch {
                Write-Log -Level WARN -Message "Could not inspect slide $slideIndex shape $shapeIndex during validation: $($_.Exception.Message)"
            }
            Release-ComObject $shape
        }
        Release-ComObject $slide
    }
    foreach ($equation in @($Manifest.equations)) {
        $expectedName = "MathType_$($equation.id)"
        if ($namedObjects -notcontains $expectedName) { $errors.Add("Missing named MathType OLE object: $expectedName") }
    }
    if ($mathTypeObjects -lt @($Manifest.equations).Count) {
        $errors.Add("Expected at least $(@($Manifest.equations).Count) MathType OLE objects; found $mathTypeObjects.")
    }
    if ($unresolvedMarkers.Count -gt 0) { $errors.Add('One or more PowerPoint equation markers remain unresolved.') }
    return [ordered]@{
        ok = $errors.Count -eq 0
        action = 'validate-pptx'
        presentation_path = $PresentationPath
        counts = [ordered]@{ mathtype_objects = $mathTypeObjects; mathml_verified = $verifiedMathML }
        named_objects = $namedObjects
        unresolved_markers = $unresolvedMarkers
        errors = @($errors)
        warnings = @('PowerPoint equations are floating OLE objects; MathType-native Word numbering and references do not exist in PPTX.')
    }
}

function Invoke-ValidatePptx {
    if ([string]::IsNullOrWhiteSpace($InputPath)) { throw 'validate-pptx requires InputPath.' }
    if ([string]::IsNullOrWhiteSpace($ManifestPath)) { throw 'validate-pptx requires ManifestPath.' }
    $input = Resolve-ExistingFile -Path $InputPath -Label 'InputPath'
    if ([IO.Path]::GetExtension($input) -ine '.pptx') { throw 'InputPath must be a .pptx file.' }
    $manifest = Read-PresentationManifest -Path $ManifestPath
    Assert-PresentationManifest -Manifest $manifest
    try {
        Start-PowerPoint -PresentationPath $input -ReadOnly
        $report = Get-PresentationValidationReport -PresentationPath $input -Manifest $manifest
    }
    finally { Stop-PowerPoint }
    Write-JsonResult $report
    if (-not $report.ok) { exit 3 }
}

function Invoke-Probe {
    param([switch]$PowerPointRequired)
    $mathTypeExe = 'C:\Program Files (x86)\MathType\MathType.exe'
    $wordTemplate = 'C:\Program Files\Microsoft Office\root\Office16\STARTUP\MathType Commands 2016.dotm'
    $progId = $null
    try { $progId = (Get-ItemProperty -LiteralPath 'Registry::HKEY_CLASSES_ROOT\Equation.DSMT4').'(default)' } catch {}
    $wordVersion = $null
    $templateLoaded = $false
    if (-not $PowerPointRequired) {
        try {
            Start-Word
            $wordVersion = $script:Word.Version
            foreach ($template in @($script:Word.Templates)) {
                if ($template.Name -like 'MathType Commands*.dotm') { $templateLoaded = $true }
                Release-ComObject $template
            }
        }
        finally { Stop-Word }
    }
    $powerPointVersion = $null
    $powerPointAddInLoaded = $false
    try {
        $script:PowerPoint = New-Object -ComObject PowerPoint.Application
        $powerPointVersion = $script:PowerPoint.Version
        foreach ($addIn in @($script:PowerPoint.AddIns)) {
            if ($addIn.Name -eq 'MathType AddIn') { $powerPointAddInLoaded = $true }
            Release-ComObject $addIn
        }
    }
    finally { Stop-PowerPoint }
    $mathTypeVersion = if (Test-Path -LiteralPath $mathTypeExe) {
        (Get-Item -LiteralPath $mathTypeExe).VersionInfo.ProductVersion
    }
    else { $null }
    $checks = [ordered]@{
        windows                    = $env:OS -eq 'Windows_NT'
        powershell                 = $PSVersionTable.PSVersion.ToString()
        mathtype_executable        = Test-Path -LiteralPath $mathTypeExe
        mathtype_executable_path   = $mathTypeExe
        mathtype_product_version   = $mathTypeVersion
        equation_dsmt4_registered  = $null -ne $progId
        equation_dsmt4_description = $progId
    }
    if ($PowerPointRequired) {
        $checks['powerpoint_com'] = $null -ne $powerPointVersion
        $checks['powerpoint_version'] = $powerPointVersion
        $checks['mathtype_powerpoint_addin_loaded'] = $powerPointAddInLoaded
    }
    else {
        $checks['word_com'] = $null -ne $wordVersion
        $checks['word_version'] = $wordVersion
        $checks['mathtype_word_template'] = Test-Path -LiteralPath $wordTemplate
        $checks['mathtype_word_template_path'] = $wordTemplate
        $checks['mathtype_template_loaded'] = $templateLoaded
    }
    $wordReady = if ($PowerPointRequired) {
        $null
    }
    else {
        $checks.windows -and $checks.word_com -and $checks.mathtype_executable -and $checks.mathtype_word_template -and $checks.mathtype_template_loaded -and $checks.equation_dsmt4_registered
    }
    $powerPointReady = if ($PowerPointRequired) {
        $checks.windows -and $checks.mathtype_executable -and $checks.equation_dsmt4_registered -and $checks.powerpoint_com -and $checks.mathtype_powerpoint_addin_loaded
    }
    else { $null }
    $ready = if ($PowerPointRequired) { $powerPointReady } else { $wordReady }
    Write-JsonResult ([ordered]@{
        ok = $ready
        action = if ($PowerPointRequired) { 'probe-pptx' } else { 'probe' }
        word_ready = $wordReady
        powerpoint_ready = $powerPointReady
        checks = $checks
    })
    if (-not $ready) { exit 2 }
}

function Invoke-ConfigureDefaults {
    if (-not (Test-Path -LiteralPath $script:PackagedDefaultsPath -PathType Leaf)) {
        throw "Packaged defaults are missing: $script:PackagedDefaultsPath"
    }
    New-Item -ItemType Directory -Path $script:UserConfigDirectory -Force | Out-Null
    Copy-Item -LiteralPath $script:PackagedDefaultsPath -Destination $script:UserDefaultsPath -Force
    if (-not (Test-Path -LiteralPath $script:WordCommandsKey)) {
        New-Item -Path $script:WordCommandsKey -Force | Out-Null
    }
    New-ItemProperty -LiteralPath $script:WordCommandsKey -Name NoEqnNumWarningDlg -PropertyType DWord -Value 0 -Force | Out-Null
    New-ItemProperty -LiteralPath $script:WordCommandsKey -Name NoInsertEqnRefDlg -PropertyType DWord -Value 1 -Force | Out-Null
    Write-JsonResult ([ordered]@{
        ok = $true
        action = 'configure-defaults'
        defaults_path = $script:UserDefaultsPath
        effective_number_format = '(1), (2), (3), ...'
        native_reference_pipeline = 'MTReference -> MTPlaceRef -> GOTOBUTTON/REF'
        scope = 'Defaults for every document processed by this bridge; warning preferences are also applied to the current MathType user profile.'
    })
}

function Invoke-Render {
    if ([string]::IsNullOrWhiteSpace($InputPath)) { throw 'render requires InputPath.' }
    if ([string]::IsNullOrWhiteSpace($OutputPath)) { throw 'render requires OutputPath.' }
    if ([string]::IsNullOrWhiteSpace($ManifestPath)) { throw 'render requires ManifestPath.' }
    $input = Resolve-ExistingFile -Path $InputPath -Label 'InputPath'
    if ([IO.Path]::GetExtension($input) -ine '.docx') { throw 'InputPath must be a .docx file.' }
    $output = Resolve-OutputFile -Path $OutputPath
    $manifest = Read-Manifest -Path $ManifestPath
    Assert-Manifest -Manifest $manifest
    $numberFields = @{}
    Set-TemporaryMathTypePreferences
    try {
        Start-Word -DocumentPath $input
        Write-Log -Level INFO -Message "Opened input DOCX: $input"
        foreach ($equation in @($manifest.equations)) {
            $numberField = Convert-Equation -Equation $equation
            if ($null -ne $numberField) {
                $numberFields[[string]$equation.id] = $numberField
                $script:HeldComObjects.Add($numberField)
            }
        }
        foreach ($reference in @($manifest.references)) {
            Insert-NativeReference -Reference $reference -TargetField $numberFields[[string]$reference.target]
        }
        $script:Document.Fields.Update() | Out-Null
        Save-DocumentAtomically -Destination $output
    }
    finally {
        Stop-Word
        Restore-MathTypePreferences
    }
    Write-JsonResult ([ordered]@{
        ok = $true
        action = 'render'
        input_path = $input
        output_path = $output
        equations = @($manifest.equations).Count
        numbered_equations = @($manifest.equations | Where-Object numbered).Count
        references = @($manifest.references).Count
        number_format = '(1), (2), (3), ...'
        reference_mechanism = 'MathType-native GOTOBUTTON/REF fields'
    })
}

function Get-ValidationReport {
    param([Parameter(Mandatory)][string]$DocumentPath, $Manifest)
    $errors = [System.Collections.Generic.List[string]]::new()
    $warnings = [System.Collections.Generic.List[string]]::new()
    $mathTypeObjects = 0
    for ($index = 1; $index -le $script:Document.InlineShapes.Count; $index++) {
        $shape = $script:Document.InlineShapes.Item($index)
        $progId = $null
        try { $progId = $shape.OLEFormat.ProgID } catch {}
        if ($progId -eq 'Equation.DSMT4') { $mathTypeObjects++ }
        Release-ComObject $shape
    }

    $numberFields = @()
    $numberValues = @()
    $referenceTargets = @()
    for ($index = 1; $index -le $script:Document.Fields.Count; $index++) {
        $field = $script:Document.Fields.Item($index)
        $code = $field.Code.Text
        if ($code -match '^\s*MACROBUTTON\s+MTPlaceRef\b') {
            $numberFields += $code
            if ($code -notmatch 'SEQ\s+MTEqn\s+\\h') { $errors.Add("Number field $index lacks hidden MTEqn sequence increment.") }
            if ($code -notmatch 'SEQ\s+MTEqn\s+\\c') { $errors.Add("Number field $index lacks current MTEqn sequence display.") }
            if ($code -match 'SEQ\s+(MTSec|MTChap)\s+\\c') { $errors.Add("Number field $index contains chapter or section display.") }
            if ($code -notmatch '\(' -or $code -notmatch '\)') { $errors.Add("Number field $index is not enclosed in parentheses.") }
            if ($code -match '\([.:-]') { $errors.Add("Number field $index contains a residual separator after the opening parenthesis.") }
            $numberParagraph = $field.Code.Paragraphs.Item(1)
            $tabCount = ([regex]::Matches($numberParagraph.Range.Text, "`t")).Count
            if ($tabCount -lt 2) { $errors.Add("Numbered display containing field $index lacks the right-alignment tab before the equation number.") }
            $hasCenterTab = $false
            $hasRightTab = $false
            for ($tabIndex = 1; $tabIndex -le $numberParagraph.TabStops.Count; $tabIndex++) {
                $tabStop = $numberParagraph.TabStops.Item($tabIndex)
                if ($tabStop.Alignment -eq 1) { $hasCenterTab = $true }
                if ($tabStop.Alignment -eq 2) { $hasRightTab = $true }
                Release-ComObject $tabStop
            }
            if (-not $hasCenterTab -or -not $hasRightTab) {
                $errors.Add("Numbered display containing field $index lacks MathType center/right tab stops.")
            }
            Release-ComObject $numberParagraph
        }
        elseif ($code -match '^\s*SEQ\s+MTEqn\s+\\c\b') {
            $parsed = 0
            if ([int]::TryParse(($field.Result.Text -replace '[^0-9]', ''), [ref]$parsed)) { $numberValues += $parsed }
        }
        elseif ($code -match '^\s*GOTOBUTTON\s+(ZEqnNum\d+)\b') {
            $referenceTargets += $Matches[1]
        }
        Release-ComObject $field
    }

    for ($position = 0; $position -lt $numberValues.Count; $position++) {
        if ($numberValues[$position] -ne ($position + 1)) {
            $errors.Add("Equation numbers are not sequential from 1: $($numberValues -join ', ').")
            break
        }
    }
    foreach ($bookmarkName in $referenceTargets) {
        if (-not $script:Document.Bookmarks.Exists($bookmarkName)) {
            $errors.Add("Reference target bookmark is missing: $bookmarkName")
        }
    }

    $plainText = $script:Document.Content.Text
    if ($plainText -match 'equation reference goes here') { $errors.Add('An unresolved MathType equation-reference placeholder remains.') }
    if ($plainText -match '\{\{(?:MATH|EQREF):') { $errors.Add('An unresolved equation or reference marker remains.') }
    if ($plainText -match 'Error! Reference source not found') { $errors.Add('Word reports a missing reference source.') }
    if ($script:Document.OMaths.Count -gt 0) { $errors.Add("Found $($script:Document.OMaths.Count) Word built-in equation object(s); MathType is required.") }

    if ($null -ne $Manifest) {
        $expectedObjects = @($Manifest.equations).Count
        $expectedNumbers = @($Manifest.equations | Where-Object numbered).Count
        $expectedReferences = @($Manifest.references).Count
        if ($mathTypeObjects -lt $expectedObjects) { $errors.Add("Expected at least $expectedObjects MathType objects; found $mathTypeObjects.") }
        if ($numberFields.Count -ne $expectedNumbers) { $errors.Add("Expected $expectedNumbers native number fields; found $($numberFields.Count).") }
        if ($referenceTargets.Count -ne $expectedReferences) { $errors.Add("Expected $expectedReferences native references; found $($referenceTargets.Count).") }
        foreach ($equation in @($Manifest.equations)) {
            if ($plainText.Contains([string]$equation.marker)) { $errors.Add("Unresolved equation marker: $($equation.marker)") }
        }
        foreach ($reference in @($Manifest.references)) {
            if ($plainText.Contains([string]$reference.marker)) { $errors.Add("Unresolved reference marker: $($reference.marker)") }
        }
    }
    else { $warnings.Add('No manifest supplied; expected counts and exact markers were not checked.') }

    return [ordered]@{
        ok = $errors.Count -eq 0
        action = 'validate'
        document_path = $DocumentPath
        counts = [ordered]@{
            mathtype_objects = $mathTypeObjects
            word_builtin_omath = $script:Document.OMaths.Count
            native_number_fields = $numberFields.Count
            native_reference_fields = $referenceTargets.Count
        }
        equation_numbers = $numberValues
        reference_bookmarks = $referenceTargets
        errors = @($errors)
        warnings = @($warnings)
    }
}

function Invoke-Validate {
    if ([string]::IsNullOrWhiteSpace($InputPath)) { throw 'validate requires InputPath.' }
    $input = Resolve-ExistingFile -Path $InputPath -Label 'InputPath'
    $manifest = $null
    if (-not [string]::IsNullOrWhiteSpace($ManifestPath)) {
        $manifest = Read-Manifest -Path $ManifestPath
        Assert-Manifest -Manifest $manifest
    }
    try {
        Start-Word -ReadOnly -DocumentPath $input
        $script:Document.Fields.Update() | Out-Null
        $report = Get-ValidationReport -DocumentPath $input -Manifest $manifest
    }
    finally { Stop-Word }
    Write-JsonResult $report
    if (-not $report.ok) { exit 3 }
}

function Invoke-Update {
    if ([string]::IsNullOrWhiteSpace($InputPath)) { throw 'update requires InputPath.' }
    $input = Resolve-ExistingFile -Path $InputPath -Label 'InputPath'
    $destination = if ([string]::IsNullOrWhiteSpace($OutputPath)) { $input } else { Resolve-OutputFile -Path $OutputPath }
    if ($destination -eq $input -and -not $Overwrite) { throw 'Updating in place requires -Overwrite.' }
    try {
        Start-Word -DocumentPath $input
        $updated = $script:Document.Fields.Update()
        Save-DocumentAtomically -Destination $destination
    }
    finally { Stop-Word }
    Write-JsonResult ([ordered]@{ ok = $true; action = 'update'; output_path = $destination; fields_update_result = $updated })
}

try {
    switch ($Action) {
        'probe' { Invoke-Probe }
        'probe-pptx' { Invoke-Probe -PowerPointRequired }
        'configure-defaults' { Invoke-ConfigureDefaults }
        'render' { Invoke-Render }
        'validate' { Invoke-Validate }
        'update' { Invoke-Update }
        'render-pptx' { Invoke-RenderPptx }
        'validate-pptx' { Invoke-ValidatePptx }
    }
}
catch {
    Write-Log -Level ERROR -Message $_.Exception.ToString()
    Write-JsonResult ([ordered]@{ ok = $false; action = $Action; error = $_.Exception.Message })
    exit 1
}
finally {
    Stop-PowerPoint
    Stop-Word
    try { Restore-MathTypePreferences } catch { Write-Log -Level WARN -Message $_.Exception.Message }
}
