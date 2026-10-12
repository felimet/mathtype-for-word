<#
Finds PowerShell 7 for the bridge and, when none exists, installs PowerShell 7.6 LTS for the current user.

This script must run under Windows PowerShell 5.1, which every supported Windows has, because it runs before
PowerShell 7 exists. It installs the official ZIP release from GitHub into %USERPROFILE%\.mathtype-for-word\PowerShell:
no administrator rights, no prompt, no PATH or registry change. The folder sits outside AppData because an
MSIX-packaged client redirects AppData writes to a private copy. The archive must match the SHA-256 pinned below,
and pwsh.exe must carry a valid Microsoft Authenticode signature.

Prints one ASCII JSON line: ok, path, version, installed, and for a new install the folder, source and sha256.
#>
[CmdletBinding()]
param(
    # Overrides for tests: another release location, version and archive hash, another install folder and lock,
    # and ignoring PowerShell 7 already present.
    [string]$DownloadBase = 'https://github.com/PowerShell/PowerShell/releases/download',
    [string]$Version,
    [string]$ExpectedSha256,
    [string]$InstallRoot = (Join-Path $env:USERPROFILE '.mathtype-for-word\PowerShell'),
    [string]$MutexName = 'Local\MathTypeForWordAgentPowerShellInstall',
    [switch]$IgnoreExisting
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue' # Windows PowerShell 5.1 renders download progress so slowly it dominates
if ($PSVersionTable.PSVersion.Major -lt 6) {
    # Below a PowerShell 7 process, Windows PowerShell inherits PowerShell 7's module folders and then cannot load
    # its own script cmdlets (Get-FileHash "is not recognized"); keep only the folders Windows PowerShell owns.
    $env:PSModulePath = (@($env:PSModulePath -split ';') |
        Where-Object { $_ -and $_ -notmatch '(?<!Windows)PowerShell\\(7[^\\]*\\)?Modules\\?$' }) -join ';'
}
# PowerShell 7.6 LTS, supported until 2028-11-14. Update the version and both hashes together, from the release's
# asset digests: gh api repos/PowerShell/PowerShell/releases/tags/v<version> --jq '.assets[] | [.name, .digest]'
$script:PinnedVersion = '7.6.6'
$script:PinnedSha256 = @{
    x64   = '02fe458be20493fbdf43f61ea20610b811ee6c738ab1676c61b9cfcd1a33c860'
    arm64 = 'bbde9dda31d148415eccb5fbe1638e6400a144187b006e5b3fd8ec2f39d781be'
}
$script:MinimumMajorVersion = 7
$script:Release = if ($Version) { $Version } else { $script:PinnedVersion }
$script:InstallWaitMilliseconds = 20 * 60 * 1000
$script:StagingPrefix = '.download-'
$script:CacheFolder = '.cache' # a verified archive kept after a failed install, so a retry does not download it again
$script:RenameAttempts = 8 # an on-access scanner can hold a just-extracted file for a moment
$script:RenameFirstDelayMilliseconds = 100
$script:MicrosoftSigner = '(^|,\s*)O=Microsoft Corporation(,|$)'

function Write-AsciiJson {
    param([Parameter(Mandatory)][System.Collections.IDictionary]$Value)
    # 5.1 writes stdout in the console code page; escaping keeps CJK user-profile paths intact for the caller.
    $json = $Value | ConvertTo-Json -Compress
    [regex]::Replace($json, '[^\x00-\x7F]', { param($match) '\u{0:x4}' -f [int][char]$match.Value })
}

function Get-PowerShellVersion {
    param([Parameter(Mandatory)][string]$Path)
    try { return [version](Get-Item -LiteralPath $Path).VersionInfo.FileVersion } catch { return $null }
}

function Find-PowerShell7 {
    # Whether a candidate is this script's own copy comes from where it was listed, never from comparing path strings:
    # an 8.3 short name and its long form are the same folder.
    $candidates = @()
    if (-not $IgnoreExisting) {
        $external = @(Get-Command pwsh.exe -CommandType Application -ErrorAction SilentlyContinue | ForEach-Object Source)
        foreach ($root in @($env:ProgramFiles, $env:ProgramW6432) | Where-Object { $_ } | Select-Object -Unique) {
            $external += Join-Path $root 'PowerShell\7\pwsh.exe'
        }
        $candidates += @($external | ForEach-Object { [pscustomobject]@{ Path = $_; Own = $false } })
    }
    if (Test-Path -LiteralPath $InstallRoot -PathType Container) {
        $installed = Get-ChildItem -LiteralPath $InstallRoot -Directory |
            Where-Object { -not $_.Name.StartsWith('.') } |
            ForEach-Object { Join-Path $_.FullName 'pwsh.exe' } |
            Sort-Object { Get-PowerShellVersion $_ } -Descending
        $candidates += @($installed | ForEach-Object { [pscustomobject]@{ Path = $_; Own = $true } })
    }
    # This script's own copies count only from the release it installs, so a plugin release with a newer pin replaces
    # them; PowerShell 7 on PATH or in Program Files is serviced by Windows and counts from 7.0.
    $ownMinimum = [version]$script:Release
    $externalMinimum = [version]::new($script:MinimumMajorVersion, 0)
    foreach ($candidate in $candidates) {
        if (-not (Test-Path -LiteralPath $candidate.Path -PathType Leaf)) { continue }
        $found = Get-PowerShellVersion $candidate.Path
        $minimum = if ($candidate.Own) { $ownMinimum } else { $externalMinimum }
        if ($null -ne $found -and $found -ge $minimum) {
            return [ordered]@{ path = $candidate.Path; version = $found.ToString(3); own = $candidate.Own }
        }
        # A Microsoft Store app-execution alias on PATH has no version resource; it is PowerShell 7 by definition.
        if ($null -eq $found -and -not $candidate.Own -and $candidate.Path -like '*\Microsoft\WindowsApps\pwsh.exe') {
            return [ordered]@{ path = $candidate.Path; version = $null; own = $false }
        }
    }
    return $null
}

function Get-WindowsArchitecture {
    if (-not [Environment]::Is64BitOperatingSystem) { throw 'Installing PowerShell 7 for this plugin needs 64-bit Windows.' }
    try { $name = "$([Runtime.InteropServices.RuntimeInformation]::OSArchitecture)" }
    catch {
        # .NET Framework before 4.7.1 lacks RuntimeInformation. Some MCP clients pass no PROCESSOR_* variables,
        # so the environment is only the fallback.
        $name = @($env:PROCESSOR_ARCHITEW6432, $env:PROCESSOR_ARCHITECTURE) | Where-Object { $_ } | Select-Object -First 1
    }
    switch -Regex ("$name") {
        '^Arm64$' { return 'arm64' }
        '^(X64|AMD64|)$' { return 'x64' }
        default { throw "PowerShell 7 has no release for this Windows architecture ($name)." }
    }
}

function Move-Directory {
    param([Parameter(Mandatory)][string]$Source, [Parameter(Mandatory)][string]$Destination)
    # One rename on one volume: a failed attempt leaves the destination absent. Move-Item in Windows PowerShell 5.1
    # falls back to moving file by file and would leave a partial copy behind.
    $delay = $script:RenameFirstDelayMilliseconds
    for ($attempt = 1; ; $attempt++) {
        try { [IO.Directory]::Move($Source, $Destination); return }
        catch [IO.IOException], [UnauthorizedAccessException] {
            if ($attempt -ge $script:RenameAttempts) { throw }
            Start-Sleep -Milliseconds $delay
            $delay *= 2
        }
    }
}

function Install-PowerShell7 {
    $architecture = Get-WindowsArchitecture
    $release = $script:Release
    $expected = if ($ExpectedSha256) { $ExpectedSha256.ToLowerInvariant() } else { $script:PinnedSha256[$architecture] }
    if ($release -ne $script:PinnedVersion -and -not $ExpectedSha256) { throw "No pinned SHA-256 for PowerShell $release." }
    $name = "PowerShell-$release-win-$architecture.zip"
    $source = "$DownloadBase/v$release/$name"

    # SystemDefault (0) already negotiates TLS 1.2 and 1.3; only an explicit older list needs TLS 1.2 added for GitHub.
    $protocols = [Net.ServicePointManager]::SecurityProtocol
    if ([int]$protocols -ne 0) { [Net.ServicePointManager]::SecurityProtocol = $protocols -bor [Net.SecurityProtocolType]::Tls12 }
    # An authenticating proxy gets the user's Windows credentials, as a browser would send them.
    if ($null -ne [Net.WebRequest]::DefaultWebProxy) {
        [Net.WebRequest]::DefaultWebProxy.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials
    }

    $staging = Join-Path $InstallRoot ($script:StagingPrefix + [guid]::NewGuid().ToString('N'))
    $cache = Join-Path $InstallRoot $script:CacheFolder
    $archive = Join-Path $cache $name
    $null = New-Item -ItemType Directory -Path $staging, $cache -Force
    try {
        if (-not (Test-Path -LiteralPath $archive -PathType Leaf) -or
            (Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash.ToLowerInvariant() -ne $expected) {
            $download = Join-Path $staging $name
            Invoke-WebRequest -Uri $source -OutFile $download -UseBasicParsing
            $actual = (Get-FileHash -LiteralPath $download -Algorithm SHA256).Hash.ToLowerInvariant()
            if ($actual -ne $expected) { throw "SHA-256 of $name is $actual; PowerShell $release publishes $expected." }
            Move-Item -LiteralPath $download -Destination $archive -Force
        }
        $actual = $expected
        $expanded = Join-Path $staging 'pwsh'
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [IO.Compression.ZipFile]::ExtractToDirectory($archive, $expanded)
        $executable = Join-Path $expanded 'pwsh.exe'
        if (-not (Test-Path -LiteralPath $executable -PathType Leaf)) { throw "$name contains no pwsh.exe." }
        $signature = Get-AuthenticodeSignature -LiteralPath $executable
        if ($signature.Status -ne 'Valid' -or $null -eq $signature.SignerCertificate -or
            $signature.SignerCertificate.Subject -notmatch $script:MicrosoftSigner) {
            throw "pwsh.exe in $name is not validly signed by Microsoft ($($signature.Status))."
        }
        $exitCode = -1
        try {
            $reported = & $executable -NoLogo -NoProfile -NonInteractive -Command '$PSVersionTable.PSVersion.ToString()'
            $exitCode = $LASTEXITCODE
        }
        catch { $reported = $_.Exception.Message }
        if ($exitCode -ne 0 -or "$reported".Trim() -ne $release) {
            throw ("The downloaded pwsh.exe did not start ($reported). Application control such as AppLocker may block " +
                'programs in the user profile; ask an administrator to install PowerShell 7.')
        }
        $target = Join-Path $InstallRoot $release
        if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target -Recurse -Force } # an unusable earlier copy
        Move-Directory -Source $expanded -Destination $target
        Remove-Item -LiteralPath $cache -Recurse -Force -ErrorAction SilentlyContinue
        # Older copies of this script's own. Renaming a copy that another client still runs fails, so that copy stays
        # whole and a later install removes it; a renamed copy is staging, which the next install sweeps if needed.
        foreach ($older in Get-ChildItem -LiteralPath $InstallRoot -Directory |
                Where-Object { -not $_.Name.StartsWith('.') -and $_.Name -ne $release }) {
            $retired = Join-Path $InstallRoot ($script:StagingPrefix + [guid]::NewGuid().ToString('N'))
            try { [IO.Directory]::Move($older.FullName, $retired) } catch { continue }
            Remove-Item -LiteralPath $retired -Recurse -Force -ErrorAction SilentlyContinue
        }
        return [ordered]@{ path = Join-Path $target 'pwsh.exe'; version = $release; folder = $target; source = $source; sha256 = $actual }
    }
    finally {
        Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue
        if ((Test-Path -LiteralPath $cache) -and -not (Get-ChildItem -LiteralPath $cache)) { Remove-Item -LiteralPath $cache }
    }
}

function Clear-InstallLeftovers {
    # A download kept for a retry, or staging of a killed install, once PowerShell 7 came from elsewhere (the MSI the
    # failure message recommends). Skipped while an installer holds the lock.
    if (-not (Test-Path -LiteralPath $InstallRoot -PathType Container)) { return }
    $leftovers = @(Get-ChildItem -LiteralPath $InstallRoot -Directory -Force |
        Where-Object { $_.Name -eq $script:CacheFolder -or $_.Name.StartsWith($script:StagingPrefix) })
    if ($leftovers.Count -eq 0) { return }
    $mutex = New-Object Threading.Mutex($false, $MutexName)
    try {
        try { $acquired = $mutex.WaitOne(0) } catch [Threading.AbandonedMutexException] { $acquired = $true }
        if (-not $acquired) { return }
        try { $leftovers | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue }
        finally { $mutex.ReleaseMutex() }
    }
    finally { $mutex.Dispose() }
}

try {
    $found = Find-PowerShell7
    if ($null -ne $found) {
        if (-not $found.own) { Clear-InstallLeftovers }
        Write-AsciiJson ([ordered]@{ ok = $true; installed = $false; path = $found.path; version = $found.version })
        exit 0
    }
    $null = New-Item -ItemType Directory -Path $InstallRoot -Force
    $mutex = New-Object Threading.Mutex($false, $MutexName)
    try {
        try { $acquired = $mutex.WaitOne($script:InstallWaitMilliseconds) }
        catch [Threading.AbandonedMutexException] { $acquired = $true } # a killed installer; its staging is swept below
        if (-not $acquired) { throw 'Another PowerShell 7 installation for mathtype-for-word did not finish in 20 minutes.' }
        try {
            $found = Find-PowerShell7 # another caller may have installed it while this one waited
            if ($null -ne $found) {
                Write-AsciiJson ([ordered]@{ ok = $true; installed = $false; path = $found.path; version = $found.version })
                exit 0
            }
            Get-ChildItem -LiteralPath $InstallRoot -Directory -Filter ($script:StagingPrefix + '*') |
                Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
            $result = [ordered]@{ ok = $true; installed = $true }
            foreach ($entry in (Install-PowerShell7).GetEnumerator()) { $result[$entry.Key] = $entry.Value }
            Write-AsciiJson $result
        }
        finally { $mutex.ReleaseMutex() }
    }
    finally { $mutex.Dispose() }
}
catch {
    Write-AsciiJson ([ordered]@{ ok = $false; error = $_.Exception.Message })
    exit 1
}
