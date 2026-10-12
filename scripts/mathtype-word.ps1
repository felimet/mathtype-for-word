[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet('probe', 'probe-pptx', 'configure-defaults', 'check-manifest', 'scan', 'mark', 'render', 'validate', 'update', 'render-pptx', 'validate-pptx', 'cleanup')]
    [string]$Action,

    [string]$InputPath,
    [string]$OutputPath,
    [string]$ManifestPath,
    [ValidatePattern('^[0-9a-fA-F]{32}$')]
    [string]$RunToken,
    [switch]$Overwrite,
    # The manifest covers only part of the document: markers outside it may remain and counts are lower bounds.
    [switch]$PartialManifest,
    # Validate the saved DOCX against this manifest instead of -ManifestPath (last chunk of a batched render).
    [string]$ValidationManifestPath,
    # Bookmark each numbered equation so a later session can reference it (every chunk of a batched render but
    # the last, which removes the bookmarks).
    [switch]$AnchorNumbers,
    # JSON list of exact text spans to replace with markers (mark).
    [string]$MarksPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
# A caller that pipes stderr (the MCP server) decodes the log as UTF-8; the ANSI code page (cp950, GBK,
# Shift-JIS) would garble CJK paths. Only the redirected stream changes: the console code page belongs to the
# caller's shared console. The JSON result line is ASCII-escaped (Write-JsonResult) for any caller code page.
if ([Console]::IsErrorRedirected) {
    $script:ErrorWriter = [IO.StreamWriter]::new([Console]::OpenStandardError(), [System.Text.UTF8Encoding]::new($false))
    $script:ErrorWriter.AutoFlush = $true
    [Console]::SetError($script:ErrorWriter)
}

$script:PluginRoot = Split-Path -Parent $PSScriptRoot
$script:PackagedDefaultsPath = Join-Path $script:PluginRoot 'config\defaults.json'
$script:UserConfigDirectory = Join-Path $env:APPDATA 'MathTypeForWordAgent'
$script:UserDefaultsPath = Join-Path $script:UserConfigDirectory 'defaults.json'
# The user's warning preferences while a render holds temporary values; a killed bridge leaves it for the next one.
$script:WarningJournalPath = Join-Path $script:UserConfigDirectory 'warning-preferences.json'
# Office processes a lock holder launched; it exists after a run only if that bridge died before its finally.
$script:OfficeJournalPath = Join-Path $script:UserConfigDirectory 'office-processes.json'
$script:LaunchedOffice = [System.Collections.Generic.List[object]]::new()
$script:WordCommandsKey = 'HKCU:\Software\Design Science\DSMT7\WordCommands'
$script:Word = $null
$script:WordOriginalAlerts = $null
$script:Document = $null
$script:Selection = $null
$script:PowerPoint = $null
$script:PowerPointOwned = $false
$script:PowerPointOriginalAlerts = $null
$script:Presentation = $null
$script:HeldComObjects = [System.Collections.Generic.List[object]]::new()
# MathType's Word macros show a modal dialog unless these read "1": first-number warning, reference prompt, and
# the slow field-update notice (shown when a whole-document update takes over 2 s, i.e. in large documents).
$script:WarningPreferenceNames = @('NoEqnNumWarningDlg', 'NoInsertEqnRefDlg', 'NoSlowUpdateEqnDlg')
$script:OriginalWarningPreferences = $null
$script:BridgeMutexName = 'Local\MathTypeForWordAgentBridge'
$script:BridgeMutex = $null
$script:BridgeLockWaitSeconds = if ($env:MATHTYPE_BRIDGE_LOCK_WAIT_SECONDS) { [int]$env:MATHTYPE_BRIDGE_LOCK_WAIT_SECONDS } else { 30 }
$script:WdLineSpaceAtLeast = 3
$script:WdLineSpaceExactly = 4
$script:WdWithInTable = 12
$script:WdRowHeightAtLeast = 1
$script:WdRowHeightExactly = 2
$script:RowHeightRepairs = [System.Collections.Generic.List[string]]::new()
# Hidden bookmarks (leading underscore) that tie an equation id to its numbered paragraph across Word sessions.
$script:AnchorPrefix = '_MTEq'
# Word draws an exactly spaced line at its nominal height; allow rounding before calling an object clipped.
$script:LineSpacingTolerancePoints = 0.5
$script:LineSpacingRepairs = [System.Collections.Generic.List[string]]::new()
$script:AutomationPollMilliseconds = 250
$script:AutomationTimeoutMilliseconds = if ($env:MATHTYPE_AUTOMATION_TIMEOUT_MS) {
    [int]$env:MATHTYPE_AUTOMATION_TIMEOUT_MS
}
else { 15000 }
$script:RunToken = if ($RunToken) { $RunToken.ToLowerInvariant() } else { [Guid]::NewGuid().ToString('N') }
$script:StaleFileCutoffHours = 24
$script:ClipboardAttempts = 3
# Hints for plain-text mathematics in a draft; the agent decides what each paragraph really holds.
$script:MathSignals = [regex]::new(@'
(?<greek>[\u0391-\u03A9\u03B1-\u03C9\u03D1\u03D5\u03F5])
|(?<operator>[=\u2212\u2248\u2260\u2261\u2264\u2265\u226A\u226B\u00B1\u2213\u00D7\u00F7\u221A\u221B\u2211\u220F\u222B\u222E\u2202\u221E\u2207\u2192\u2190\u2194\u21D2\u21D4\u2208\u2209\u2282\u2283\u2286\u2287\u222A\u2229\u2200\u2203\u221D\u22A5\u2225])
|(?<script>[\u2070-\u209F\u00B2\u00B3\u00B9])
|(?<tex>\\[A-Za-z]+|[A-Za-z0-9)\]}][_^]\{?[A-Za-z0-9])
|(?<function>(?<![A-Za-z])(?:sin|cos|tan|cot|sec|csc|arcsin|arccos|arctan|sinh|cosh|tanh|log|ln|lg|exp|max|min|lim|det)\s*\()
|(?<reference>(?:(?<![A-Za-z])(?:Eqs?\.|Equations?)|\u516C\u5F0F|\u5F0F)\s*[\uFF08(]\s*\d+(?:[.-]\d+)?\s*[)\uFF09]
  |(?:\u7B2C\s*)?[\uFF08(]\s*\d+(?:[.-]\d+)?\s*[)\uFF09]\s*\u5F0F)
'@, 'IgnorePatternWhitespace')
$script:MarkerPattern = '^\{\{(?:MATH|EQREF):[^{}\r\n]+\}\}$'
# Word's Find accepts at most 255 characters.
$script:FindTextLimit = 255
$script:ClipboardRetryMilliseconds = 500

if (-not ('MathTypeOleData' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Runtime.InteropServices;
public static class MathTypeOleData {
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern uint RegisterClipboardFormat(string format);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GlobalLock(IntPtr memory);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GlobalUnlock(IntPtr memory);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern UIntPtr GlobalSize(IntPtr memory);

    [DllImport("ole32.dll")]
    private static extern void ReleaseStgMedium(ref STGMEDIUM medium);

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

if (-not ('OfficeDesktop' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class OfficeDesktop {
    private delegate bool WindowVisitor(IntPtr window, IntPtr parameter);
    [DllImport("user32.dll")] private static extern bool EnumWindows(WindowVisitor visitor, IntPtr parameter);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int size);
    [DllImport("user32.dll", SetLastError = true)] private static extern bool OpenClipboard(IntPtr owner);
    [DllImport("user32.dll")] private static extern bool CloseClipboard();
    [DllImport("user32.dll")] private static extern IntPtr GetOpenClipboardWindow();
    [DllImport("user32.dll")] private static extern IntPtr GetClipboardOwner();

    // A visible Word or PowerPoint main frame, or a PowerPoint slide show (a .ppsx opened into the instance shows
    // only screenClass), ignoring dialogs: a MathType dialog in a hidden instance is what hangs it (the same rule
    // as the MCP watchdog).
    public static bool HasVisibleFrame(int processId) {
        bool found = false;
        EnumWindows(delegate (IntPtr window, IntPtr parameter) {
            uint owner;
            GetWindowThreadProcessId(window, out owner);
            if (owner != (uint)processId || !IsWindowVisible(window)) return true;
            StringBuilder name = new StringBuilder(64);
            GetClassName(window, name, name.Capacity);
            string windowClass = name.ToString();
            found = windowClass == "OpusApp" || windowClass == "PPTFrameClass" || windowClass == "screenClass";
            return !found;
        }, IntPtr.Zero);
        return found;
    }

    // Null when the clipboard can be opened; otherwise who holds it open. Opening and closing without emptying
    // leaves the contents untouched. A holder that opened it without a window cannot be named; the last writer is
    // then a clue, unless it is Word or PowerPoint, which here is usually this bridge's own instance.
    public static string DescribeClipboardBlocker() {
        if (OpenClipboard(IntPtr.Zero)) { CloseClipboard(); return null; }
        string holder = Describe(GetOpenClipboardWindow());
        if (holder != null) return "held open by " + holder;
        string writer = Describe(GetClipboardOwner());
        bool office = writer != null && (writer.StartsWith("WINWORD ") || writer.StartsWith("POWERPNT "));
        return "held open by an unidentified program (often a clipboard manager or chat app)" +
            (writer != null && !office ? "; the last program that wrote to it is " + writer : "");
    }

    private static string Describe(IntPtr window) {
        if (window == IntPtr.Zero) return null;
        uint processId;
        if (GetWindowThreadProcessId(window, out processId) == 0 || processId == 0) return null;
        try { return System.Diagnostics.Process.GetProcessById((int)processId).ProcessName + " (PID " + processId + ")"; }
        catch (ArgumentException) { return null; }
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
    [Console]::Error.WriteLine("[$([DateTime]::Now.ToString('s'))] [$Level] $Message")
}

function Enter-BridgeLock {
    # One editing job at a time: concurrent bridges race on the MathType warning preferences, the clipboard
    # used for PowerPoint, and MathType's own state.
    $script:BridgeMutex = [Threading.Mutex]::new($false, $script:BridgeMutexName)
    try { $acquired = $script:BridgeMutex.WaitOne([TimeSpan]::FromSeconds($script:BridgeLockWaitSeconds)) }
    catch [Threading.AbandonedMutexException] { $acquired = $true } # the previous holder died; the lock is ours
    if (-not $acquired) {
        $script:BridgeMutex.Dispose()
        $script:BridgeMutex = $null
        throw "Another MathType Office job is still running after $($script:BridgeLockWaitSeconds) s; retry when it has finished."
    }
    Stop-OrphanedOffice
    Restore-JournaledPreferences
}

function Exit-BridgeLock {
    if ($null -eq $script:BridgeMutex) { return }
    # Reached after Stop-PowerPoint and Stop-Word: nothing journaled is left for the next holder to end.
    Remove-Item -LiteralPath $script:OfficeJournalPath -ErrorAction SilentlyContinue
    $script:BridgeMutex.ReleaseMutex()
    $script:BridgeMutex.Dispose()
    $script:BridgeMutex = $null
}

function Add-LaunchedOffice {
    param([Parameter(Mandatory)][Diagnostics.Process]$Process)
    # Only a lock holder journals. A bridge killed with its server (job object) never reaches its finally, and the
    # Office it launched through DCOM is outside that job; the next lock holder ends it (Stop-OrphanedOffice).
    if ($null -eq $script:BridgeMutex) { return }
    $script:LaunchedOffice.Add([ordered]@{ pid = $Process.Id; start = $Process.StartTime.ToFileTimeUtc() })
    New-Item -ItemType Directory -Path $script:UserConfigDirectory -Force | Out-Null
    ConvertTo-Json -InputObject @($script:LaunchedOffice) -Compress | Set-Content -LiteralPath $script:OfficeJournalPath -Encoding utf8
}

function Stop-OrphanedOffice {
    # Runs under the bridge lock. A failure here only leaves a hidden process running; it must never block the job.
    if (-not (Test-Path -LiteralPath $script:OfficeJournalPath)) { return }
    $entries = @()
    try { $entries = @(Get-Content -Raw -LiteralPath $script:OfficeJournalPath | ConvertFrom-Json | Where-Object { $null -ne $_ }) }
    catch { Write-Log -Level WARN -Message "Ignored an unreadable Office process journal: $($_.Exception.Message)" }
    foreach ($entry in $entries) {
        try {
            $process = Get-Process -Id ([int]$entry.pid) -ErrorAction SilentlyContinue
            # Another start time means a reused PID; a visible main frame means the user took the instance over.
            if ($null -eq $process -or $process.StartTime.ToFileTimeUtc() -ne [long]$entry.start) { continue }
            if ([OfficeDesktop]::HasVisibleFrame($process.Id)) {
                Write-Log -Level WARN -Message "Left $($process.ProcessName) $($process.Id) from an interrupted run running: it shows a window."
                continue
            }
            Stop-Process -Id $process.Id -Force
            Write-Log -Level WARN -Message "Ended hidden $($process.ProcessName) $($process.Id) left by an interrupted run."
        }
        catch { Write-Log -Level WARN -Message "Could not end an Office process left by an interrupted run: $($_.Exception.Message)" }
    }
    Remove-Item -LiteralPath $script:OfficeJournalPath -Force
}

function Write-JsonResult {
    param([Parameter(Mandatory)]$Value)
    # ASCII escapes keep the JSON line valid for callers that decode with any code page.
    $Value | ConvertTo-Json -Depth 20 -Compress -EscapeHandling EscapeNonAscii
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

function Get-OfficeTemporaryPath {
    param([Parameter(Mandatory)][string]$Destination)
    $fullPath = [IO.Path]::GetFullPath($Destination)
    $extension = [IO.Path]::GetExtension($fullPath).ToLowerInvariant()
    if ($extension -notin @('.docx', '.pptx')) {
        throw "Temporary Office output must use .docx or .pptx: $fullPath"
    }
    $name = [IO.Path]::GetFileNameWithoutExtension($fullPath)
    return Join-Path (Split-Path -Parent $fullPath) ".$name.$($script:RunToken).tmp$extension"
}

function Remove-OfficeTemporaryArtifacts {
    param(
        [Parameter(Mandatory)][string]$Destination,
        [switch]$IncludeCurrentRun
    )
    $fullPath = [IO.Path]::GetFullPath($Destination)
    $extension = [IO.Path]::GetExtension($fullPath).ToLowerInvariant()
    if ($extension -notin @('.docx', '.pptx')) {
        throw "Cleanup target must use .docx or .pptx: $fullPath"
    }
    $parent = Split-Path -Parent $fullPath
    $currentPath = Get-OfficeTemporaryPath -Destination $fullPath
    $removed = [System.Collections.Generic.List[string]]::new()
    if (Test-Path -LiteralPath $parent -PathType Container) {
        $name = [regex]::Escape([IO.Path]::GetFileNameWithoutExtension($fullPath))
        $suffix = [regex]::Escape(".tmp$extension")
        $pattern = "^\.$name\.[0-9a-f]{32}$suffix$"
        $cutoff = [DateTime]::UtcNow.AddHours(-$script:StaleFileCutoffHours)
        foreach ($file in @(Get-ChildItem -LiteralPath $parent -File -Force)) {
            if ($file.Name -notmatch $pattern -or ($file.Attributes -band [IO.FileAttributes]::ReparsePoint)) { continue }
            $isCurrent = [string]::Equals($file.FullName, $currentPath, [StringComparison]::OrdinalIgnoreCase)
            if (-not (($IncludeCurrentRun -and $isCurrent) -or $file.LastWriteTimeUtc -lt $cutoff)) { continue }
            Remove-Item -LiteralPath $file.FullName -Force
            if (Test-Path -LiteralPath $file.FullName) { throw "Temporary file cleanup could not be verified: $($file.FullName)" }
            $removed.Add($file.FullName)
            Write-Log -Level INFO -Message "Removed temporary Office file: $($file.FullName)"
        }
    }
    return [ordered]@{
        removed = @($removed)
        current_run_removed = $removed.Contains($currentPath)
        current_run_remaining = Test-Path -LiteralPath $currentPath
        stale_cutoff_hours = $script:StaleFileCutoffHours
    }
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
    # Log the PID before Open/Toggle TeX, where hangs happen: the MCP timeout watchdog terminates only logged PIDs.
    $newProcesses = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id })
    if ($newProcesses.Count -eq 1) {
        Write-Log -Level INFO -Message "WORD_PID=$($newProcesses[0].Id)"
        Add-LaunchedOffice $newProcesses[0]
    }
    elseif ($newProcesses.Count -gt 1) { Write-Log -Level WARN -Message "Multiple new Word PIDs observed: $($newProcesses.Id -join ', ')" }
    $script:Word.Visible = $false
    $script:WordOriginalAlerts = $script:Word.DisplayAlerts
    $script:Word.DisplayAlerts = 0
    if ([string]::IsNullOrWhiteSpace($DocumentPath)) {
        $script:Document = $script:Word.Documents.Add()
    }
    else {
        $script:Document = $script:Word.Documents.Open($DocumentPath, $false, [bool]$ReadOnly)
    }
    $script:Selection = $script:Word.Selection
}

function Connect-PowerPoint {
    # PowerPoint is single-instance: COM attaches to the user's PowerPoint when one is running. Only an instance
    # this bridge launched may be quit, and an attached instance gets its alert setting back.
    if ($null -ne $script:PowerPoint) { return }
    $before = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $script:PowerPoint = New-Object -ComObject PowerPoint.Application
    $launched = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id })
    $script:PowerPointOwned = $launched.Count -eq 1
    if ($script:PowerPointOwned) {
        Write-Log -Level INFO -Message "POWERPOINT_PID=$($launched[0].Id)"
        Add-LaunchedOffice $launched[0]
    }
    else { Write-Log -Level INFO -Message 'Attached to a PowerPoint instance that was already running; it will be left running.' }
    $script:PowerPointOriginalAlerts = $script:PowerPoint.DisplayAlerts
    $script:PowerPoint.DisplayAlerts = 1
}

function Start-PowerPoint {
    param([Parameter(Mandatory)][string]$PresentationPath, [switch]$ReadOnly)
    Connect-PowerPoint
    # In the user's instance, Open may hand back a presentation the user is editing, and closing it would discard
    # unsaved work. Refuse instead of touching it.
    $target = [IO.Path]::GetFullPath($PresentationPath)
    $targetName = [IO.Path]::GetFileName($target)
    foreach ($open in @($script:PowerPoint.Presentations)) {
        $openPath = [string]$open.FullName
        $openName = [string]$open.Name
        Release-ComObject $open
        # A OneDrive/SharePoint-synced file reports an https FullName instead of the local path that opened it.
        # ponytail: cloud files match by name only, so another open cloud deck of the same name is refused too.
        $sameCloudFile = $openPath -match '^https?://' -and [string]::Equals($openName, $targetName, [StringComparison]::OrdinalIgnoreCase)
        if ($sameCloudFile -or [string]::Equals($openPath, $target, [StringComparison]::OrdinalIgnoreCase)) {
            throw "The presentation is already open in PowerPoint; save and close it, then retry: $target"
        }
    }
    $script:Presentation = $script:PowerPoint.Presentations.Open(
        $PresentationPath,
        [bool]$ReadOnly,
        $false,
        $false
    )
}

function Stop-Word {
    param([switch]$Save)
    if ($null -ne $script:Document) {
        try { $script:Document.Close($(if ($Save) { -1 } else { 0 })) } catch {
            Write-Log -Level WARN -Message "Could not close Word document cleanly: $($_.Exception.Message)"
        }
    }
    if ($null -ne $script:Word) {
        $userWindows = $false
        try {
            # A file the user opens while this hidden automation instance runs can land in it and make it visible.
            $remaining = [int]$script:Word.Documents.Count + [int]$script:Word.ProtectedViewWindows.Count
            $userWindows = $remaining -gt 0 -and [bool]$script:Word.Visible
        }
        catch { Write-Log -Level WARN -Message "Could not inspect Word before quitting: $($_.Exception.Message)" }
        if ($userWindows) {
            try { $script:Word.DisplayAlerts = $script:WordOriginalAlerts } catch {}
            Write-Log -Level WARN -Message 'Left Word running: the user opened a document in the bridge instance.'
        }
        else {
            # Only hidden bridge documents can remain; never let a save prompt decide their fate.
            try { $script:Word.Quit(0) } catch {
                Write-Log -Level WARN -Message "Could not quit Word cleanly: $($_.Exception.Message)"
            }
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

function Close-Presentation {
    # A failed Close keeps the reference and throws: the caller sees the real cause, and Stop-PowerPoint knows the
    # bridge's own presentation is still open rather than counting it as the user's.
    if ($null -eq $script:Presentation) { return }
    $script:Presentation.Close()
    Release-ComObject $script:Presentation
    $script:Presentation = $null
}

function Stop-PowerPoint {
    $ownStillOpen = 0
    try { Close-Presentation } catch {
        Write-Log -Level WARN -Message "Could not close PowerPoint presentation cleanly: $($_.Exception.Message)"
        $ownStillOpen = 1
    }
    if ($null -ne $script:PowerPoint) {
        $userPresentations = $null
        try {
            $userPresentations = [int]$script:PowerPoint.Presentations.Count + [int]$script:PowerPoint.ProtectedViewWindows.Count - $ownStillOpen
        }
        catch { Write-Log -Level WARN -Message "Could not count open presentations; PowerPoint is left running: $($_.Exception.Message)" }
        if ($script:PowerPointOwned -and $userPresentations -eq 0) {
            try { $script:PowerPoint.Quit() } catch {
                Write-Log -Level WARN -Message "Could not quit PowerPoint cleanly: $($_.Exception.Message)"
            }
        }
        else {
            # The user's instance, or ours after the user opened a presentation in it: never quit it.
            try { $script:PowerPoint.DisplayAlerts = $script:PowerPointOriginalAlerts } catch {}
            Write-Log -Level INFO -Message 'Left PowerPoint running; only the bridge presentation was closed.'
        }
    }
    Release-ComObject $script:Presentation
    $script:Presentation = $null
    Release-ComObject $script:PowerPoint
    $script:PowerPoint = $null
    $script:PowerPointOwned = $false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Set-WarningPreference {
    # MathType's Word add-in reads these preferences as REG_SZ "0"/"1" (MathType writes them that way); a
    # REG_DWORD value is ignored, and the warning dialog then blocks the hidden Word instance.
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][AllowEmptyString()][string]$Value)
    New-ItemProperty -LiteralPath $script:WordCommandsKey -Name $Name -PropertyType String -Value $Value -Force | Out-Null
}

function Set-WarningValues {
    # $null removes a value that did not exist before the run.
    param([Parameter(Mandatory)][System.Collections.IDictionary]$Values)
    foreach ($name in $script:WarningPreferenceNames) {
        $value = $Values[$name]
        if ($null -eq $value) { Remove-ItemProperty -LiteralPath $script:WordCommandsKey -Name $name -ErrorAction SilentlyContinue }
        else { Set-WarningPreference -Name $name -Value ([string]$value) }
    }
}

function Restore-JournaledPreferences {
    # Runs under the bridge lock: a bridge killed before its finally left the user's values in the journal, and
    # they must be back before anyone takes a new snapshot.
    if (-not (Test-Path -LiteralPath $script:WarningJournalPath) -or -not (Test-Path -LiteralPath $script:WordCommandsKey)) { return }
    $journal = $null
    $problem = 'not a JSON object'
    try { $journal = Get-Content -Raw -LiteralPath $script:WarningJournalPath | ConvertFrom-Json -AsHashtable }
    catch { $problem = $_.Exception.Message }
    if ($journal -isnot [System.Collections.IDictionary]) {
        # A crash can leave the journal empty or NUL-filled. The user's values are lost, but a journal that can
        # never be read must not block every later job.
        $quarantine = "$($script:WarningJournalPath).corrupt-$([DateTime]::Now.ToString('yyyyMMddHHmmssfff'))"
        Move-Item -LiteralPath $script:WarningJournalPath -Destination $quarantine -Force
        Write-Log -Level WARN -Message "Unreadable warning-preference journal ($problem) moved to ${quarantine}; MathType warning preferences keep their current values."
        return
    }
    Set-WarningValues -Values $journal
    Remove-Item -LiteralPath $script:WarningJournalPath -Force
    Write-Log -Level WARN -Message 'Restored MathType warning preferences left by an interrupted run.'
}

function Set-TemporaryMathTypePreferences {
    if (-not (Test-Path -LiteralPath $script:WordCommandsKey)) {
        throw 'MathType 7 WordCommands registry key was not found.'
    }
    $key = Get-Item -LiteralPath $script:WordCommandsKey
    try {
        $original = [ordered]@{}
        foreach ($name in $script:WarningPreferenceNames) { $original[$name] = $key.GetValue($name) } # $null when absent
    }
    finally { $key.Close() }
    # Journal first: if this process is killed, the next bridge (or the MCP watchdog) restores from it.
    New-Item -ItemType Directory -Path (Split-Path -Parent $script:WarningJournalPath) -Force | Out-Null
    $temporary = "$($script:WarningJournalPath).tmp"
    $original | ConvertTo-Json -Compress | Set-Content -LiteralPath $temporary -Encoding utf8
    Move-Item -LiteralPath $temporary -Destination $script:WarningJournalPath -Force
    $script:OriginalWarningPreferences = $original
    foreach ($name in $script:WarningPreferenceNames) { Set-WarningPreference -Name $name -Value '1' }
}

function Restore-MathTypePreferences {
    if ($null -eq $script:OriginalWarningPreferences -or -not (Test-Path -LiteralPath $script:WordCommandsKey)) { return }
    Set-WarningValues -Values $script:OriginalWarningPreferences
    Remove-Item -LiteralPath $script:WarningJournalPath -ErrorAction SilentlyContinue
    $script:OriginalWarningPreferences = $null
}

function Add-MissingProperties {
    # StrictMode makes reading an absent JSON property throw. Give every expected property a value once, so
    # optional fields take their defaults and missing required fields reach the validation messages below.
    param([Parameter(Mandatory)]$InputObject, [Parameter(Mandatory)][hashtable]$Defaults)
    foreach ($name in $Defaults.Keys) {
        if ($null -eq $InputObject.PSObject.Properties[$name]) {
            $InputObject | Add-Member -NotePropertyName $name -NotePropertyValue $Defaults[$name]
        }
    }
}

function Read-Manifest {
    param([Parameter(Mandatory)][string]$Path)
    $resolved = Resolve-ExistingFile -Path $Path -Label 'ManifestPath'
    # -NoEnumerate keeps a one-element top-level array an array, so it is rejected instead of unwrapped.
    $manifest = Get-Content -Raw -LiteralPath $resolved | ConvertFrom-Json -NoEnumerate
    if ($manifest -isnot [System.Management.Automation.PSCustomObject]) { throw 'Manifest must be a JSON object.' }
    Add-MissingProperties -InputObject $manifest -Defaults @{
        schema_version = $null; equations = $null; references = $null; exact_line_spacing = 'at_least'; external_targets = @()
    }
    if ($manifest.schema_version -ne 1) { throw 'Manifest schema_version must be 1.' }
    if ($manifest.equations -isnot [array]) { throw 'Manifest must contain an equations array.' }
    if ($null -eq $manifest.references) { $manifest.references = @() }
    elseif ($manifest.references -isnot [array]) { throw 'Manifest references must be an array.' }
    if ($manifest.external_targets -isnot [array]) { throw 'Manifest external_targets must be an array.' }
    foreach ($equation in $manifest.equations) {
        if ($equation -isnot [System.Management.Automation.PSCustomObject]) { throw 'Every manifest equation must be a JSON object.' }
        Add-MissingProperties -InputObject $equation -Defaults @{ id = $null; marker = $null; tex = $null; layout = $null; numbered = $false }
    }
    foreach ($reference in $manifest.references) {
        if ($reference -isnot [System.Management.Automation.PSCustomObject]) { throw 'Every manifest reference must be a JSON object.' }
        Add-MissingProperties -InputObject $reference -Defaults @{ marker = $null; target = $null }
    }
    return $manifest
}

function Assert-Manifest {
    param([Parameter(Mandatory)]$Manifest)
    $ids = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    # Invoke-Render keys number fields in a case-insensitive hashtable: "E1" and "e1" would share one field.
    $foldedIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $markers = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($equation in @($Manifest.equations)) {
        # Ids are compared as strings: a numeric id would match a string target only after coercion.
        if ($equation.id -isnot [string] -or [string]::IsNullOrWhiteSpace($equation.id)) { throw 'Every equation requires a non-empty string id.' }
        if (-not $ids.Add([string]$equation.id)) { throw "Duplicate equation id: $($equation.id)" }
        if (-not $foldedIds.Add([string]$equation.id)) { throw "Equation id differs from another only by case: $($equation.id)" }
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
    if (@('at_least', 'keep') -notcontains [string]$Manifest.exact_line_spacing) {
        throw 'exact_line_spacing must be "at_least" or "keep".'
    }
    # Numbered equations of an earlier session in a batched render, reachable through their anchor bookmarks.
    $external = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($id in @($Manifest.external_targets)) {
        if ($id -isnot [string] -or $ids.Contains($id) -or -not $external.Add($id)) { throw "Invalid external target: $id" }
    }
    foreach ($reference in @($Manifest.references)) {
        if ([string]::IsNullOrWhiteSpace($reference.marker)) { throw 'Every reference requires a marker.' }
        if (-not $markers.Add([string]$reference.marker)) { throw "Duplicate marker: $($reference.marker)" }
        if ($reference.target -isnot [string] -or -not ($ids.Contains($reference.target) -or $external.Contains($reference.target))) {
            throw "Reference target does not exist: $($reference.target)"
        }
        if ($external.Contains($reference.target)) { continue }
        $target = @($Manifest.equations | Where-Object id -eq $reference.target)[0]
        if (-not [bool]$target.numbered) { throw "Reference target is not numbered: $($reference.target)" }
    }
}

function Read-PresentationManifest {
    param([Parameter(Mandatory)][string]$Path)
    $resolved = Resolve-ExistingFile -Path $Path -Label 'ManifestPath'
    $manifest = Get-Content -Raw -LiteralPath $resolved | ConvertFrom-Json -NoEnumerate
    if ($manifest -isnot [System.Management.Automation.PSCustomObject]) { throw 'Presentation manifest must be a JSON object.' }
    Add-MissingProperties -InputObject $manifest -Defaults @{ schema_version = $null; equations = $null }
    if ($manifest.schema_version -ne 1) { throw 'Presentation manifest schema_version must be 1.' }
    if ($manifest.equations -isnot [array]) { throw 'Presentation manifest must contain an equations array.' }
    foreach ($equation in $manifest.equations) {
        if ($equation -isnot [System.Management.Automation.PSCustomObject]) { throw 'Every presentation equation must be a JSON object.' }
        Add-MissingProperties -InputObject $equation -Defaults @{ id = $null; marker = $null; tex = $null; mathml = $null; height_points = $null }
    }
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
        if ([string]::IsNullOrWhiteSpace($equation.tex)) { throw "Presentation equation '$($equation.id)' requires TeX for silent conversion." }
        if ([string]$equation.tex -match '[\r\n]') { throw "Presentation equation '$($equation.id)' TeX must be one line." }
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
    # Count in the story text instead of re-running Find from each match: inside table cells Word's Find can
    # return the same match again, which made the old search loop spin forever. Markers never overlap.
    $occurrences = [regex]::Matches($script:Document.Content.Text, [regex]::Escape($Marker)).Count
    if ($occurrences -ne 1) {
        throw "Marker must occur exactly once; found ${occurrences}: $Marker"
    }
    $search = $script:Document.Content.Duplicate
    $find = $search.Find
    try {
        $find.ClearFormatting()
        $find.Text = $Marker
        $find.Forward = $true
        $find.Wrap = 0
        $find.MatchWildcards = $false
        $find.MatchWholeWord = $false
        # Ids are case-sensitive and the count above is too; a case-insensitive Find could pick another marker.
        $find.MatchCase = $true
        if (-not $find.Execute() -or [string]$search.Text -cne $Marker) {
            throw "Word Find could not locate the marker that occurs in the text: $Marker"
        }
        return $script:Document.Range($search.Start, $search.End)
    }
    finally {
        Release-ComObject $find
        Release-ComObject $search
    }
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

function Get-ClippingLineSpacing {
    # Returns the exact line spacing (pt) of the object's paragraph when it is lower than the object, else $null.
    param([Parameter(Mandatory)]$Shape)
    $format = $Shape.Range.ParagraphFormat
    try {
        if ($format.LineSpacingRule -ne $script:WdLineSpaceExactly) { return $null }
        $spacing = [double]$format.LineSpacing
        if ([double]$Shape.Height -gt $spacing + $script:LineSpacingTolerancePoints) { return $spacing }
        return $null
    }
    finally { Release-ComObject $format }
}

function Repair-ClippedLineSpacing {
    # Exact line spacing clips an object taller than the line. "At least" with the same value keeps the
    # paragraph's spacing on every line the equation does not need to grow.
    param([Parameter(Mandatory)]$Shape, [Parameter(Mandatory)][string]$EquationId)
    $spacing = Get-ClippingLineSpacing -Shape $Shape
    if ($null -eq $spacing) { return }
    $format = $Shape.Range.ParagraphFormat
    try {
        $format.LineSpacingRule = $script:WdLineSpaceAtLeast
        $format.LineSpacing = $spacing
        # Exact spacing ignored the document line grid; at-least would snap every line to it again.
        $format.DisableLineHeightGrid = -1
    }
    finally { Release-ComObject $format }
    $script:LineSpacingRepairs.Add($EquationId)
    Write-Log -Level INFO -Message "Changed exact $spacing pt line spacing to at-least for '$EquationId' ($([Math]::Round($Shape.Height, 1)) pt tall)."
}

function Get-ClippingRowHeight {
    # Returns the exact height (pt) of the object's table row when it is lower than the object itself needs, else
    # $null; throws when the row cannot be inspected. Cell, not Rows.Item: a table with vertically merged cells
    # exposes no single row (Word error 5991), but a cell still reports its row's rule and height.
    # ponytail: counts only the object, its cell padding and its paragraph spacing; other lines in the cell can
    # still clip it unnoticed. Rendering does not rely on this estimate (Repair-ClippedRowHeight).
    param([Parameter(Mandatory)]$Shape)
    if (-not [bool]$Shape.Range.Information($script:WdWithInTable)) { return $null }
    $cell = $Shape.Range.Cells.Item(1)
    $format = $Shape.Range.ParagraphFormat
    try {
        if ($cell.HeightRule -ne $script:WdRowHeightExactly) { return $null }
        $needed = [double]$Shape.Height + [double]$cell.TopPadding + [double]$cell.BottomPadding +
            [double]$format.SpaceBefore + [double]$format.SpaceAfter
        if ($needed -gt [double]$cell.Height + $script:LineSpacingTolerancePoints) { return [double]$cell.Height }
        return $null
    }
    finally {
        Release-ComObject $cell
        Release-ComObject $format
    }
}

function Repair-ClippedRowHeight {
    # An at-least row of the same height looks exactly like the exact row while the cell content fits and grows
    # only when it does not, so every exact row that receives an equation switches without estimating the content.
    # ponytail: only the innermost row; an exact outer row of a nested table is left as it is.
    param([Parameter(Mandatory)]$Shape, [Parameter(Mandatory)][string]$EquationId)
    if (-not [bool]$Shape.Range.Information($script:WdWithInTable)) { return }
    $cell = $null
    try {
        $cell = $Shape.Range.Cells.Item(1)
        if ($cell.HeightRule -ne $script:WdRowHeightExactly) { return }
        $height = [double]$cell.Height
        $cell.HeightRule = $script:WdRowHeightAtLeast # a cell's rule applies to its whole row
    }
    catch {
        Write-Log -Level WARN -Message "Could not inspect the table row of '$EquationId'; validation reports it: $($_.Exception.Message)"
        return
    }
    finally { Release-ComObject $cell }
    $script:RowHeightRepairs.Add($EquationId)
    Write-Log -Level INFO -Message "Changed exact $height pt table row height to at-least for '$EquationId' ($([Math]::Round($Shape.Height, 1)) pt tall)."
}

function Get-AnchorName {
    param([Parameter(Mandatory)][string]$Id)
    # Ids are free-form strings; a bookmark name allows 40 letters, digits and underscores. (SHA256.HashData and
    # Convert.ToHexString would need .NET 5, which PowerShell 7.0 lacks.)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $digest = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Id)) }
    finally { $sha.Dispose() }
    return $script:AnchorPrefix + [BitConverter]::ToString($digest, 0, 16).Replace('-', '').ToLowerInvariant()
}

function Find-AnchoredNumberField {
    # A reference target numbered in an earlier session: its paragraph carries the anchor bookmark.
    param([Parameter(Mandatory)][string]$Id)
    $name = Get-AnchorName -Id $Id
    if (-not $script:Document.Bookmarks.Exists($name)) {
        throw "Reference target '$Id' was not numbered in an earlier session (anchor $name is missing)."
    }
    $bookmark = $script:Document.Bookmarks.Item($name)
    $fields = $bookmark.Range.Fields
    $found = $null
    $count = 0
    try {
        for ($index = 1; $index -le $fields.Count; $index++) {
            $field = $fields.Item($index)
            if ($field.Code.Text -match '^\s*MACROBUTTON\s+MTPlaceRef\b') {
                $count++
                if ($null -eq $found) { $found = $field; continue }
            }
            Release-ComObject $field
        }
    }
    finally {
        Release-ComObject $fields
        Release-ComObject $bookmark
    }
    if ($count -gt 1) {
        # The anchor covers the whole paragraph: with two numbers in it, the reference could land on the wrong one.
        Release-ComObject $found
        throw "Reference target '$Id' shares its paragraph with $($count - 1) other MathType number field(s); put each numbered display equation in its own paragraph."
    }
    if ($null -eq $found) { throw "Reference target '$Id' has no MathType number field in its anchored paragraph." }
    return $found
}

function Remove-Anchors {
    $bookmarks = $script:Document.Bookmarks
    $showHidden = $bookmarks.ShowHidden
    $bookmarks.ShowHidden = $true
    $removed = 0
    try {
        for ($index = $bookmarks.Count; $index -ge 1; $index--) {
            $bookmark = $bookmarks.Item($index)
            if ($bookmark.Name.StartsWith($script:AnchorPrefix, [StringComparison]::Ordinal)) {
                $bookmark.Delete()
                $removed++
            }
            Release-ComObject $bookmark
        }
    }
    finally {
        $bookmarks.ShowHidden = $showHidden
        Release-ComObject $bookmarks
    }
    if ($removed -gt 0) { Write-Log -Level INFO -Message "Removed $removed cross-session anchor bookmark(s)." }
}

function Get-AnchorCount {
    $bookmarks = $script:Document.Bookmarks
    $showHidden = $bookmarks.ShowHidden
    $bookmarks.ShowHidden = $true
    try {
        $count = 0
        for ($index = 1; $index -le $bookmarks.Count; $index++) {
            $bookmark = $bookmarks.Item($index)
            if ($bookmark.Name.StartsWith($script:AnchorPrefix, [StringComparison]::Ordinal)) { $count++ }
            Release-ComObject $bookmark
        }
        return $count
    }
    finally {
        $bookmarks.ShowHidden = $showHidden
        Release-ComObject $bookmarks
    }
}

function Convert-Equation {
    param([Parameter(Mandatory)]$Equation, [switch]$RepairExactHeights, [switch]$AnchorNumber)
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
    if ($RepairExactHeights) {
        Repair-ClippedLineSpacing -Shape $shape -EquationId ([string]$Equation.id)
        Repair-ClippedRowHeight -Shape $shape -EquationId ([string]$Equation.id)
    }

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
        if ($AnchorNumber) {
            $paragraph = $shape.Range.Paragraphs.Item(1)
            $null = $script:Document.Bookmarks.Add((Get-AnchorName -Id ([string]$Equation.id)), $paragraph.Range)
            Release-ComObject $paragraph
        }
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
    # -Markers: a marked draft, checked for each marker exactly once instead of the render validation.
    param([Parameter(Mandatory)][string]$Destination, $Manifest, [string[]]$Markers)
    $temporary = Get-OfficeTemporaryPath -Destination $Destination
    $null = Remove-OfficeTemporaryArtifacts -Destination $Destination
    try {
        Write-Log -Level INFO -Message "Saving temporary DOCX: $temporary"
        $script:Document.SaveAs2($temporary, 16)
        Write-Log -Level INFO -Message 'Word SaveAs2 returned.'
        Stop-Word
        try {
            Start-Word -ReadOnly -DocumentPath $temporary
            if ($Markers) { $validation = Get-MarkerReport -Markers $Markers }
            else {
                $script:Document.Fields.Update() | Out-Null
                $validation = Get-ValidationReport -DocumentPath $Destination -Manifest $Manifest
            }
        }
        finally { Stop-Word }
        if (-not $validation.ok) {
            throw "Generated DOCX validation failed: $($validation.errors -join '; ')"
        }
        Move-Item -LiteralPath $temporary -Destination $Destination -Force
        if (-not (Test-Path -LiteralPath $Destination -PathType Leaf) -or (Get-Item -LiteralPath $Destination).Length -le 0) {
            throw "Published DOCX could not be verified: $Destination"
        }
        Write-Log -Level INFO -Message "Moved completed DOCX to: $Destination"
        return $validation
    }
    finally {
        $null = Remove-OfficeTemporaryArtifacts -Destination $Destination -IncludeCurrentRun
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
    $shapeRange = $null
    $wordShape = $null
    $oleObject = $null
    try {
        $delimited = '\[' + [string]$Equation.tex + '\]'
        $script:Document.Content.Text = $delimited
        $sourceRange = $script:Document.Range(0, $delimited.Length)
        $sourceRange.Select()
        $before = $script:Document.InlineShapes.Count
        $null = $script:Word.Run('MTCommand_TeXToggle')
        Release-ComObject $sourceRange
        if ($script:Document.InlineShapes.Count -ne ($before + 1)) {
            throw 'Hidden Word conversion did not create exactly one MathType object.'
        }
        $wordShape = $script:Document.InlineShapes.Item($script:Document.InlineShapes.Count)
        # The clipboard carries the OLE object to PowerPoint. A clipboard manager may hold it open for a moment;
        # a program that never closes it is named in the error, since only the user can close that program.
        for ($attempt = 1; $null -eq $shapeRange; $attempt++) {
            try {
                $wordShape.Range.Copy()
                $shapeRange = $Slide.Shapes.Paste()
            }
            catch {
                if ($attempt -ge $script:ClipboardAttempts) {
                    $blocker = [OfficeDesktop]::DescribeClipboardBlocker()
                    if ($null -eq $blocker) { throw }
                    throw "The Windows clipboard is $blocker; close or restart the program holding it, then retry. ($($_.Exception.Message))"
                }
                Write-Log -Level WARN -Message "Clipboard transfer attempt $attempt failed; retrying: $($_.Exception.Message)"
                Start-Sleep -Milliseconds $script:ClipboardRetryMilliseconds
            }
        }
        for ($index = 1; $index -le $shapeRange.Count; $index++) {
            $candidate = $shapeRange.Item($index)
            $progId = $null
            try { $progId = $candidate.OLEFormat.ProgID } catch {}
            if ($progId -eq 'Equation.DSMT4' -and $null -eq $shape) {
                $shape = $candidate
                continue
            }
            $candidate.Delete()
            Release-ComObject $candidate
        }
        if ($null -eq $shape) { throw 'Pasted PowerPoint shapes contain no Equation.DSMT4 object.' }
        $expectedSignature = Get-MathMLTextSignature -MathML ([string]$Equation.mathml)
        $oleObject = $shape.OLEFormat.Object
        $embeddedMathML = [MathTypeOleData]::GetMathML($oleObject)
        if ($embeddedMathML -notmatch '<(?:[A-Za-z_][\w.-]*:)?math\b') {
            throw 'Embedded MathType object did not expose a MathML root after silent conversion.'
        }
        if ($embeddedMathML.Contains([char]0xFFFD) -or $embeddedMathML -match '&#x0*FFFD;') {
            throw 'Embedded MathType object contains a Unicode replacement character.'
        }
        $actualSignature = Get-MathMLTextSignature -MathML $embeddedMathML
        if ([string]::IsNullOrWhiteSpace($actualSignature) -or $actualSignature -ne $expectedSignature) {
            throw "MathML text signature mismatch: expected '$expectedSignature', got '$actualSignature'."
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
        throw "Silent PowerPoint MathType OLE insertion failed for '$($Equation.id)': $($_.Exception.Message)"
    }
    finally {
        Release-ComObject $shapeRange
        Release-ComObject $wordShape
        Release-ComObject $oleObject
        if ($null -ne $script:Document) { $script:Document.Content.Text = '' }
    }
}

function Save-PresentationAtomically {
    param([Parameter(Mandatory)][string]$Destination, [Parameter(Mandatory)]$Manifest)
    $temporary = Get-OfficeTemporaryPath -Destination $Destination
    $null = Remove-OfficeTemporaryArtifacts -Destination $Destination
    try {
        Write-Log -Level INFO -Message "Saving temporary PPTX: $temporary"
        $script:Presentation.SaveAs($temporary, 24)
        # Reopen in the same instance: quitting and relaunching races PowerPoint's asynchronous exit.
        Close-Presentation
        try {
            Start-PowerPoint -PresentationPath $temporary -ReadOnly
            $validation = Get-PresentationValidationReport -PresentationPath $Destination -Manifest $Manifest
        }
        finally {
            # Keep a validation error visible; a presentation left open is retried and counted when PowerPoint stops.
            try { Close-Presentation } catch { Write-Log -Level WARN -Message "Could not close the validated PPTX: $($_.Exception.Message)" }
        }
        if (-not $validation.ok) {
            throw "Generated PPTX validation failed: $($validation.errors -join '; ')"
        }
        Move-Item -LiteralPath $temporary -Destination $Destination -Force
        if (-not (Test-Path -LiteralPath $Destination -PathType Leaf) -or (Get-Item -LiteralPath $Destination).Length -le 0) {
            throw "Published PPTX could not be verified: $Destination"
        }
        return $validation
    }
    finally {
        $null = Remove-OfficeTemporaryArtifacts -Destination $Destination -IncludeCurrentRun
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
        Start-Word
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
        $validation = Save-PresentationAtomically -Destination $output -Manifest $manifest
    }
    finally {
        Stop-PowerPoint
        Stop-Word
    }
    Write-JsonResult ([ordered]@{
        ok = $true
        action = 'render-pptx'
        input_path = $input
        output_path = $output
        equations = @($manifest.equations).Count
        object_type = 'Equation.DSMT4 floating OLE'
        numbering_and_references = 'Not available through the MathType 7 PowerPoint integration.'
        validation = $validation
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
    try {
        Start-Word
        $wordVersion = $script:Word.Version
        foreach ($template in @($script:Word.Templates)) {
            if ($template.Name -like 'MathType Commands*.dotm') { $templateLoaded = $true }
            Release-ComObject $template
        }
    }
    finally { Stop-Word }
    $powerPointVersion = $null
    $powerPointAddInLoaded = $false
    if ($PowerPointRequired) {
        try {
            Connect-PowerPoint
            $powerPointVersion = $script:PowerPoint.Version
            foreach ($addIn in @($script:PowerPoint.AddIns)) {
                if ($addIn.Name -eq 'MathType AddIn') { $powerPointAddInLoaded = $true }
                Release-ComObject $addIn
            }
        }
        finally { Stop-PowerPoint }
    }
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
        word_com                   = $null -ne $wordVersion
        word_version               = $wordVersion
        mathtype_word_template     = Test-Path -LiteralPath $wordTemplate
        mathtype_word_template_path = $wordTemplate
        mathtype_template_loaded   = $templateLoaded
    }
    if ($PowerPointRequired) {
        $checks['powerpoint_com'] = $null -ne $powerPointVersion
        $checks['powerpoint_version'] = $powerPointVersion
        $checks['mathtype_powerpoint_addin_loaded'] = $powerPointAddInLoaded
    }
    $wordReady = $checks.windows -and $checks.word_com -and $checks.mathtype_executable -and $checks.mathtype_word_template -and $checks.mathtype_template_loaded -and $checks.equation_dsmt4_registered
    $powerPointReady = if ($PowerPointRequired) {
        $wordReady -and $checks.powerpoint_com -and $checks.mathtype_powerpoint_addin_loaded
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
    Set-WarningPreference -Name NoEqnNumWarningDlg -Value '0'
    Set-WarningPreference -Name NoInsertEqnRefDlg -Value '1'
    Write-JsonResult ([ordered]@{
        ok = $true
        action = 'configure-defaults'
        defaults_path = $script:UserDefaultsPath
        effective_number_format = '(1), (2), (3), ...'
        native_reference_pipeline = 'MTReference -> MTPlaceRef -> GOTOBUTTON/REF'
        scope = 'Defaults for every document processed by this bridge; warning preferences are also applied to the current MathType user profile.'
    })
}

function Invoke-CheckManifest {
    # Office-free: lets a batched render reject a bad manifest before its first Word session. The normalized copy
    # is the only manifest the batch planner reads, so it never parses the file differently from the bridge.
    if ([string]::IsNullOrWhiteSpace($ManifestPath)) { throw 'check-manifest requires ManifestPath.' }
    $manifest = Read-Manifest -Path $ManifestPath
    Assert-Manifest -Manifest $manifest
    $equations = @(foreach ($equation in $manifest.equations) {
            [ordered]@{
                id = [string]$equation.id; marker = [string]$equation.marker; tex = [string]$equation.tex
                layout = [string]$equation.layout; numbered = [bool]$equation.numbered
            }
        })
    $references = @(foreach ($reference in $manifest.references) {
            [ordered]@{ marker = [string]$reference.marker; target = [string]$reference.target }
        })
    Write-JsonResult ([ordered]@{
        ok = $true
        action = 'check-manifest'
        manifest_path = (Resolve-Path -LiteralPath $ManifestPath).Path
        equations = $equations.Count
        numbered_equations = @($equations | Where-Object { $_.numbered }).Count
        references = $references.Count
        exact_line_spacing = [string]$manifest.exact_line_spacing
        manifest = [ordered]@{
            schema_version = 1
            exact_line_spacing = [string]$manifest.exact_line_spacing
            equations = $equations
            references = $references
        }
    })
}

function Get-ParagraphSpans {
    # Start and end of every main-story paragraph (table cells included), in Word's own numbering from 1.
    $spans = [System.Collections.Generic.List[object]]::new()
    foreach ($paragraph in $script:Document.Paragraphs) {
        $range = $paragraph.Range
        $spans.Add([pscustomobject]@{ start = $range.Start; end = $range.End; range = $range })
        Release-ComObject $paragraph
    }
    return , $spans
}

function Find-ParagraphIndex {
    # Binary search: 1-based number of the paragraph that holds $Position, or 0.
    param([Parameter(Mandatory)]$Spans, [Parameter(Mandatory)][int]$Position)
    $low = 0
    $high = $Spans.Count - 1
    while ($low -le $high) {
        $middle = [int][Math]::Floor(($low + $high) / 2)
        if ($Position -lt $Spans[$middle].start) { $high = $middle - 1 }
        elseif ($Position -ge $Spans[$middle].end) { $low = $middle + 1 }
        else { return $middle + 1 }
    }
    return 0
}

function Get-ScriptRuns {
    # Word subscript or superscript runs; plain text loses them (x2 typed with a superscript means x^2).
    # One bounded Find per paragraph, without its paragraph or end-of-cell mark: across a cell mark Word's Find can
    # return an earlier match, and a stall there must not end the scan of the rest of the document.
    param(
        [Parameter(Mandatory)][ValidateSet('subscript', 'superscript')][string]$Kind,
        [Parameter(Mandatory)][AllowEmptyCollection()]$Spans,
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.HashSet[int]]$Incomplete
    )
    $runs = [System.Collections.Generic.List[object]]::new()
    for ($index = 1; $index -le $Spans.Count; $index++) {
        $span = $Spans[$index - 1]
        $limit = $span.end - 1
        if ($limit -le $span.start) { continue }
        $font = $span.range.Font
        $flag = if ($Kind -eq 'subscript') { $font.Subscript } else { $font.Superscript }
        Release-ComObject $font
        if ($flag -eq 0) { continue } # 0 none, -1 the whole paragraph, 9999999 part of it
        $range = $script:Document.Range($span.start, $limit)
        $find = $range.Find
        $find.ClearFormatting()
        $find.Text = ''
        $find.Format = $true
        $find.Forward = $true
        $find.Wrap = 0
        if ($Kind -eq 'subscript') { $find.Font.Subscript = -1 } else { $find.Font.Superscript = -1 }
        $last = -1
        # A collapsed range would search on to the end of the document, so stop at the limit first.
        while ($range.Start -lt $limit -and $find.Execute()) {
            if ($range.Start -le $last -or $range.End -gt $limit -or $range.End -le $range.Start) {
                [void]$Incomplete.Add($index)
                break
            }
            $last = $range.Start
            # Offset in the paragraph's text string (fields and shapes make Word positions differ from it).
            $prefix = $script:Document.Range($span.start, $range.Start)
            $runs.Add([pscustomobject]@{
                paragraph = $index
                kind = $Kind
                text = [string]$range.Text
                offset = ([string]$prefix.Text).Length
                note_reference = ([int]$range.Footnotes.Count + [int]$range.Endnotes.Count) -gt 0
            })
            Release-ComObject $prefix
            $range.SetRange($range.End, $limit)
        }
        Release-ComObject $find
        Release-ComObject $range
    }
    return $runs.ToArray()
}

function Invoke-Scan {
    if ([string]::IsNullOrWhiteSpace($InputPath)) { throw 'scan requires InputPath.' }
    if ([string]::IsNullOrWhiteSpace($OutputPath)) { throw 'scan requires OutputPath (the candidates JSON).' }
    $input = Resolve-ExistingFile -Path $InputPath -Label 'InputPath'
    if ([IO.Path]::GetExtension($input) -ine '.docx') { throw 'InputPath must be a .docx file.' }
    $output = Resolve-OutputFile -Path $OutputPath -Extension '.json'
    try {
        Start-Word -ReadOnly -DocumentPath $input
        $spans = Get-ParagraphSpans
        $found = @{}
        $incomplete = [System.Collections.Generic.HashSet[int]]::new()
        $runs = @(Get-ScriptRuns -Kind subscript -Spans $spans -Incomplete $incomplete) +
            @(Get-ScriptRuns -Kind superscript -Spans $spans -Incomplete $incomplete)
        foreach ($run in @($runs | Sort-Object paragraph, offset)) {
            if (-not $found.ContainsKey($run.paragraph)) { $found[$run.paragraph] = [System.Collections.Generic.List[object]]::new() }
            $found[$run.paragraph].Add([ordered]@{ kind = $run.kind; text = $run.text; offset = $run.offset; note_reference = $run.note_reference })
        }
        $equations = @{}
        for ($index = 1; $index -le $script:Document.OMaths.Count; $index++) {
            $omath = $script:Document.OMaths.Item($index)
            $paragraph = Find-ParagraphIndex -Spans $spans -Position $omath.Range.Start
            if (-not $equations.ContainsKey($paragraph)) { $equations[$paragraph] = [System.Collections.Generic.List[string]]::new() }
            $equations[$paragraph].Add([string]$omath.Range.Text)
            Release-ComObject $omath
        }
        $paragraphs = [System.Collections.Generic.List[object]]::new()
        for ($index = 1; $index -le $spans.Count; $index++) {
            $range = $spans[$index - 1].range
            $text = ([string]$range.Text).TrimEnd("`r", "`a")
            $signals = @(foreach ($match in $script:MathSignals.Matches($text)) {
                    $kind = @('greek', 'operator', 'script', 'tex', 'function', 'reference') |
                        Where-Object { $match.Groups[$_].Success } | Select-Object -First 1
                    [ordered]@{ kind = $kind; text = $match.Value; offset = $match.Index }
                })
            # @() keeps one-element and empty lists as JSON arrays.
            $formatted = @(if ($found.ContainsKey($index)) { $found[$index] })
            $wordEquations = @(if ($equations.ContainsKey($index)) { $equations[$index] })
            if ($signals.Count + $formatted.Count + $wordEquations.Count -eq 0 -and -not $incomplete.Contains($index)) { continue }
            $paragraphs.Add([ordered]@{
                paragraph = $index
                in_table = [bool]$range.Information($script:WdWithInTable)
                text = $text
                signals = $signals
                formatted = $formatted
                word_equations = $wordEquations
            })
        }
        $plainText = $script:Document.Content.Text
        $mathTypeObjects = 0
        for ($index = 1; $index -le $script:Document.InlineShapes.Count; $index++) {
            $shape = $script:Document.InlineShapes.Item($index)
            try { if ($shape.OLEFormat.ProgID -eq 'Equation.DSMT4') { $mathTypeObjects++ } } catch {}
            Release-ComObject $shape
        }
        $candidates = [ordered]@{
            schema = 'mathtype-for-word-scan/1'
            document_path = $input
            paragraphs_total = $spans.Count
            existing_markers = [regex]::Matches($plainText, '\{\{(?:MATH|EQREF):').Count
            existing_mathtype_objects = $mathTypeObjects
            word_equations = $script:Document.OMaths.Count
            # Paragraphs whose sub/superscripts Word's Find could not list completely; read them in Word.
            formatted_incomplete = @($incomplete | Sort-Object)
            paragraphs = @($paragraphs)
        }
        foreach ($span in $spans) { Release-ComObject $span.range }
    }
    finally { Stop-Word }
    $temporary = "$output.$($script:RunToken).tmp"
    [IO.File]::WriteAllText($temporary, ($candidates | ConvertTo-Json -Depth 8), [Text.UTF8Encoding]::new($false))
    Move-Item -LiteralPath $temporary -Destination $output -Force
    Write-JsonResult ([ordered]@{
        ok = $true
        action = 'scan'
        input_path = $input
        candidates_path = $output
        paragraphs_with_candidates = $paragraphs.Count
        word_equations = $candidates.word_equations
        existing_markers = $candidates.existing_markers
        existing_mathtype_objects = $mathTypeObjects
    })
}

function Read-Marks {
    param([Parameter(Mandatory)][string]$Path)
    $resolved = Resolve-ExistingFile -Path $Path -Label 'MarksPath'
    $document = Get-Content -Raw -LiteralPath $resolved | ConvertFrom-Json -NoEnumerate
    if ($document -isnot [System.Management.Automation.PSCustomObject]) { throw 'Marks file must be a JSON object.' }
    Add-MissingProperties -InputObject $document -Defaults @{ schema_version = $null; marks = $null }
    if ($document.schema_version -ne 1) { throw 'Marks file schema_version must be 1.' }
    if ($document.marks -isnot [array] -or $document.marks.Count -eq 0) { throw 'Marks file must contain a non-empty marks array.' }
    $markers = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($mark in $document.marks) {
        if ($mark -isnot [System.Management.Automation.PSCustomObject]) { throw 'Every mark must be a JSON object.' }
        Add-MissingProperties -InputObject $mark -Defaults @{ paragraph = $null; text = $null; occurrence = 1; marker = $null }
        $label = "Mark '$($mark.marker)'"
        if (-not ($mark.paragraph -is [long] -or $mark.paragraph -is [int]) -or $mark.paragraph -lt 1) { throw "$label requires a paragraph number from 1." }
        if (-not ($mark.occurrence -is [long] -or $mark.occurrence -is [int]) -or $mark.occurrence -lt 1) { throw "$label requires an occurrence from 1." }
        # Tab and manual line break (\v) may go; paragraph marks, note references (\u0002), objects (\u0001) and page,
        # column or section breaks are document structure that a replacement would delete with the formula.
        if ($mark.text -isnot [string] -or $mark.text.Length -eq 0 -or $mark.text -match '[\x00-\x08\x0A\x0C-\x1F]') {
            throw "$label requires plain text within one paragraph, without paragraph marks, note references, objects or breaks."
        }
        if ($mark.text.Replace('^', '^^').Length -gt $script:FindTextLimit) {
            throw "$label text is longer than $($script:FindTextLimit) characters (each ^ counts twice); mark a shorter unique part."
        }
        if ($mark.marker -isnot [string] -or $mark.marker -notmatch $script:MarkerPattern) {
            throw "Every marker must read {{MATH:id}} or {{EQREF:id}}: $($mark.marker)"
        }
        if (-not $markers.Add($mark.marker)) { throw "Duplicate marker: $($mark.marker)" }
    }
    return @($document.marks)
}

function Find-TextInRange {
    # The start and end of the given exact occurrence of $Text between $Start and $End, or $null.
    param([int]$Start, [int]$End, [string]$Text, [int]$Occurrence)
    $range = $script:Document.Range($Start, $End)
    $find = $range.Find
    $find.ClearFormatting()
    $find.Text = $Text.Replace('^', '^^') # ^ starts a special character even without wildcards
    $find.MatchCase = $true
    $find.MatchWholeWord = $false
    $find.MatchWildcards = $false
    $find.Forward = $true
    $find.Wrap = 0
    $find.Format = $false
    $count = 0
    $last = -1
    try {
        while ($find.Execute() -and $range.Start -gt $last -and $range.End -le $End) {
            $last = $range.Start
            # Find also matches other widths and case forms of some characters; count exact text only.
            if ([string]$range.Text -ceq $Text) {
                $count++
                if ($count -eq $Occurrence) { return @($range.Start, $range.End) }
            }
            $range.SetRange($range.End, $End)
        }
        return $null
    }
    finally {
        Release-ComObject $find
        Release-ComObject $range
    }
}

function Get-MarkerReport {
    param([Parameter(Mandatory)][string[]]$Markers)
    $plainText = $script:Document.Content.Text
    $errors = [System.Collections.Generic.List[string]]::new()
    foreach ($marker in $Markers) {
        $count = [regex]::Matches($plainText, [regex]::Escape($marker)).Count
        if ($count -ne 1) { $errors.Add("Marker $marker occurs $count times.") }
    }
    return [ordered]@{ ok = $errors.Count -eq 0; markers = $Markers.Count; errors = @($errors) }
}

function Invoke-Mark {
    if ([string]::IsNullOrWhiteSpace($InputPath)) { throw 'mark requires InputPath.' }
    if ([string]::IsNullOrWhiteSpace($OutputPath)) { throw 'mark requires OutputPath.' }
    if ([string]::IsNullOrWhiteSpace($MarksPath)) { throw 'mark requires MarksPath.' }
    $input = Resolve-ExistingFile -Path $InputPath -Label 'InputPath'
    if ([IO.Path]::GetExtension($input) -ine '.docx') { throw 'InputPath must be a .docx file.' }
    $output = Resolve-OutputFile -Path $OutputPath
    if ([string]::Equals($output, $input, [StringComparison]::OrdinalIgnoreCase)) { throw 'mark writes a new DOCX; OutputPath must differ from InputPath.' }
    $marks = Read-Marks -Path $MarksPath
    try {
        # Read-only: the source is never saved, and a copy the user has open in Word does not block this.
        Start-Word -ReadOnly -DocumentPath $input
        $plainText = $script:Document.Content.Text
        foreach ($mark in $marks) {
            if ($plainText.Contains([string]$mark.marker)) { throw "The document already contains $($mark.marker)." }
        }
        $spans = Get-ParagraphSpans
        # Locate every span before changing anything, so positions refer to the unedited document.
        $targets = foreach ($mark in $marks) {
            if ($mark.paragraph -gt $spans.Count) { throw "Paragraph $($mark.paragraph) does not exist; the document has $($spans.Count)." }
            $span = $spans[[int]$mark.paragraph - 1]
            $position = Find-TextInRange -Start $span.start -End $span.end -Text $mark.text -Occurrence ([int]$mark.occurrence)
            if ($null -eq $position) {
                throw "Paragraph $($mark.paragraph) has no occurrence $($mark.occurrence) of '$($mark.text)'."
            }
            $probe = $script:Document.Range($position[0], $position[1])
            try { $embedded = [int]$probe.OMaths.Count + [int]$probe.Fields.Count + [int]$probe.InlineShapes.Count }
            finally { Release-ComObject $probe }
            if ($embedded -gt 0) {
                throw "Mark $($mark.marker) covers a Word equation, field or object; mark plain text only."
            }
            [pscustomobject]@{ start = $position[0]; end = $position[1]; marker = [string]$mark.marker }
        }
        foreach ($span in $spans) { Release-ComObject $span.range }
        $ordered = @($targets | Sort-Object start)
        for ($index = 1; $index -lt $ordered.Count; $index++) {
            if ($ordered[$index].start -lt $ordered[$index - 1].end) {
                throw "Marks $($ordered[$index - 1].marker) and $($ordered[$index].marker) overlap."
            }
        }
        # Last first: replacing a span never moves one that is still to be replaced.
        foreach ($target in @($ordered | Sort-Object start -Descending)) {
            $range = $script:Document.Range($target.start, $target.end)
            $range.Text = $target.marker
            $range.Font.Subscript = 0
            $range.Font.Superscript = 0
            Release-ComObject $range
        }
        $validation = Save-DocumentAtomically -Destination $output -Markers @($marks | ForEach-Object { [string]$_.marker })
    }
    finally { Stop-Word }
    Write-JsonResult ([ordered]@{
        ok = $true
        action = 'mark'
        input_path = $input
        output_path = $output
        marks = $marks.Count
        validation = $validation
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
    $validationManifest = $manifest
    if (-not [string]::IsNullOrWhiteSpace($ValidationManifestPath)) {
        $validationManifest = Read-Manifest -Path $ValidationManifestPath
        Assert-Manifest -Manifest $validationManifest
    }
    $repairExactHeights = [string]$manifest.exact_line_spacing -eq 'at_least'
    $numberFields = @{}
    Set-TemporaryMathTypePreferences
    try {
        Start-Word -DocumentPath $input
        Write-Log -Level INFO -Message "Opened input DOCX: $input"
        foreach ($equation in @($manifest.equations)) {
            $numberField = Convert-Equation -Equation $equation -RepairExactHeights:$repairExactHeights -AnchorNumber:$AnchorNumbers
            if ($null -ne $numberField) {
                $numberFields[[string]$equation.id] = $numberField
                $script:HeldComObjects.Add($numberField)
            }
        }
        foreach ($reference in @($manifest.references)) {
            $target = [string]$reference.target
            if (-not $numberFields.ContainsKey($target)) {
                $numberFields[$target] = Find-AnchoredNumberField -Id $target
                $script:HeldComObjects.Add($numberFields[$target])
            }
            Insert-NativeReference -Reference $reference -TargetField $numberFields[$target]
        }
        # The last session leaves no helper bookmarks in the user's document.
        if (-not $AnchorNumbers) { Remove-Anchors }
        $script:Document.Fields.Update() | Out-Null
        $validation = Save-DocumentAtomically -Destination $output -Manifest $validationManifest
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
        line_spacing_changed_for = @($script:LineSpacingRepairs)
        row_height_changed_for = @($script:RowHeightRepairs)
        validation = $validation
    })
}

function Get-ValidationReport {
    param([Parameter(Mandatory)][string]$DocumentPath, $Manifest)
    $errors = [System.Collections.Generic.List[string]]::new()
    $warnings = [System.Collections.Generic.List[string]]::new()
    $mathTypeObjects = 0
    $clippedObjects = 0
    $clippedRows = 0
    $uninspectedRows = 0
    for ($index = 1; $index -le $script:Document.InlineShapes.Count; $index++) {
        $shape = $script:Document.InlineShapes.Item($index)
        $progId = $null
        try { $progId = $shape.OLEFormat.ProgID } catch {}
        if ($progId -eq 'Equation.DSMT4') {
            $mathTypeObjects++
            if ($null -ne (Get-ClippingLineSpacing -Shape $shape)) { $clippedObjects++ }
            try { if ($null -ne (Get-ClippingRowHeight -Shape $shape)) { $clippedRows++ } }
            catch { $uninspectedRows++ }
        }
        Release-ComObject $shape
    }
    if ($clippedObjects -gt 0) {
        $warnings.Add("$clippedObjects MathType object(s) are taller than the exact line spacing of their paragraph and are clipped; use at-least line spacing there.")
    }
    if ($clippedRows -gt 0) {
        $warnings.Add("$clippedRows MathType object(s) are taller than the exact height of their table row and are clipped; use an at-least row height there.")
    }
    if ($uninspectedRows -gt 0) {
        $warnings.Add("$uninspectedRows MathType object(s) sit in table rows whose height could not be inspected; check them for clipping.")
    }
    $anchors = Get-AnchorCount
    if ($anchors -gt 0 -and -not $AnchorNumbers) {
        $errors.Add("$anchors cross-session anchor bookmark(s) of an unfinished batched render remain.")
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
    $leftoverMarkers = [regex]::Matches($plainText, '\{\{(?:MATH|EQREF):').Count
    if ($leftoverMarkers -gt 0) {
        $message = "$leftoverMarkers unresolved equation or reference marker(s) remain."
        if ($PartialManifest) { $warnings.Add("$message They are outside this partial manifest.") } else { $errors.Add($message) }
    }
    if ($plainText -match 'Error! Reference source not found') { $errors.Add('Word reports a missing reference source.') }
    if ($script:Document.OMaths.Count -gt 0) { $errors.Add("Found $($script:Document.OMaths.Count) Word built-in equation object(s); MathType is required.") }

    if ($null -ne $Manifest) {
        $expectedObjects = @($Manifest.equations).Count
        $expectedNumbers = @($Manifest.equations | Where-Object numbered).Count
        $expectedReferences = @($Manifest.references).Count
        if ($mathTypeObjects -lt $expectedObjects) { $errors.Add("Expected at least $expectedObjects MathType objects; found $mathTypeObjects.") }
        if ($PartialManifest) {
            # Equations and references outside a partial manifest (earlier batches) may already exist.
            if ($numberFields.Count -lt $expectedNumbers) { $errors.Add("Expected at least $expectedNumbers native number fields; found $($numberFields.Count).") }
            if ($referenceTargets.Count -lt $expectedReferences) { $errors.Add("Expected at least $expectedReferences native references; found $($referenceTargets.Count).") }
        }
        else {
            if ($numberFields.Count -ne $expectedNumbers) { $errors.Add("Expected $expectedNumbers native number fields; found $($numberFields.Count).") }
            if ($referenceTargets.Count -ne $expectedReferences) { $errors.Add("Expected $expectedReferences native references; found $($referenceTargets.Count).") }
        }
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
        $validation = Save-DocumentAtomically -Destination $destination -Manifest $null
    }
    finally { Stop-Word }
    Write-JsonResult ([ordered]@{ ok = $true; action = 'update'; output_path = $destination; fields_update_result = $updated; validation = $validation })
}

function Invoke-Cleanup {
    if ([string]::IsNullOrWhiteSpace($OutputPath)) { throw 'cleanup requires OutputPath.' }
    $cleanup = Remove-OfficeTemporaryArtifacts -Destination $OutputPath -IncludeCurrentRun
    Write-JsonResult ([ordered]@{ ok = -not $cleanup.current_run_remaining; action = 'cleanup'; cleanup = $cleanup })
    if ($cleanup.current_run_remaining) { exit 4 }
}

try {
    if (@('render', 'render-pptx', 'configure-defaults') -contains $Action) { Enter-BridgeLock }
    switch ($Action) {
        'probe' { Invoke-Probe }
        'probe-pptx' { Invoke-Probe -PowerPointRequired }
        'configure-defaults' { Invoke-ConfigureDefaults }
        'check-manifest' { Invoke-CheckManifest }
        'scan' { Invoke-Scan }
        'mark' { Invoke-Mark }
        'render' { Invoke-Render }
        'validate' { Invoke-Validate }
        'update' { Invoke-Update }
        'render-pptx' { Invoke-RenderPptx }
        'validate-pptx' { Invoke-ValidatePptx }
        'cleanup' { Invoke-Cleanup }
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
    Exit-BridgeLock
}
