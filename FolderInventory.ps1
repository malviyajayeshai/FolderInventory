<#
.SYNOPSIS
    FolderInventory - scans a Windows drive or folder and writes a self-contained,
    interactive HTML inventory report (file table, folder tree, filters, CSV export).

.DESCRIPTION
    Recursively enumerates a drive or directory using a bounded pool of PowerShell
    runspaces (no background jobs, no PowerShell 7-only features). Every folder is
    queued exactly once on a shared work queue; workers pull folders from that queue,
    list their immediate contents, and push one result per folder. The main thread
    collects the results, links folders into a tree, computes folder totals in a single
    bottom-up pass, and writes one HTML file that contains the CSS, JavaScript and data.

    Nothing on disk is modified, deleted, or has its permissions changed. Folders that
    cannot be read are recorded with the reason, and the report is marked Partial.

    Compatible with Windows PowerShell 5.1 (Windows 10/11) and PowerShell 7+.

.PARAMETER RootPath
    Drive or folder to scan. Default: the current file-system directory.
    Position 0.

.PARAMETER OutputHtml
    Path of the HTML report to write. Default: FolderInventory.html in the current
    file-system directory. If an existing folder is given, FolderInventory.html is
    written inside it. Missing parent folders are created. Position 1.

.PARAMETER ThrottleLimit
    Maximum number of concurrent scanning workers (1-64). Default: 4.
    Use 1-2 for spinning hard drives and network shares, 4-8 for SSDs. Position 2.

.PARAMETER IncludeHidden
    Include files and folders that have the Hidden attribute.

.PARAMETER IncludeSystem
    Include files and folders that have the System attribute (where accessible).
    An item that is both Hidden and System needs both switches.

.PARAMETER FollowReparsePoints
    Descend into directory junctions and directory symbolic links. Off by default,
    because links can point outside the selected root or create loops; a depth limit
    of 256 levels is enforced either way. Links that are not followed are listed in
    the report.

.PARAMETER TimeoutMinutes
    Optional overall scan time limit. 0 (default) means no limit. When reached, the
    scan stops, everything collected so far is reported, unscanned folders are listed,
    and the report is marked Partial.

.EXAMPLE
    & "$HOME\Downloads\FolderInventory.ps1" -RootPath "D:\" -OutputHtml "$HOME\Downloads\FolderInventory.html" -ThrottleLimit 4

.EXAMPLE
    & ".\FolderInventory.ps1"

.EXAMPLE
    & ".\FolderInventory.ps1" "E:\Photos" "$HOME\Desktop\Photos inventory.html" 2

.NOTES
    Exit codes: 0 = complete scan, report written
                2 = partial scan (some folders/files could not be read), report written
                1 = fatal error, no report written
#>
[CmdletBinding(PositionalBinding = $false)]
param(
    [Parameter(Position = 0)]
    [string]$RootPath = '',

    [Parameter(Position = 1)]
    [string]$OutputHtml = '',

    [Parameter(Position = 2)]
    [ValidateRange(1, 64)]
    [int]$ThrottleLimit = 4,

    [Parameter()]
    [switch]$IncludeHidden,

    [Parameter()]
    [switch]$IncludeSystem,

    [Parameter()]
    [switch]$FollowReparsePoints,

    [Parameter()]
    [ValidateRange(0, 10080)]
    [int]$TimeoutMinutes = 0
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:ToolVersion = '1.0.0'
$script:MaxDepth = 256
$script:IsWindowsFileSystem = ([System.IO.Path]::DirectorySeparatorChar -eq '\')
$script:Invariant = [System.Globalization.CultureInfo]::InvariantCulture

# ---------------------------------------------------------------------------
# Extension -> category mapping (documented in README). Extensions are compared
# lower-case and without the leading dot. Anything not listed is "Other";
# files without an extension are "No extension".
# ---------------------------------------------------------------------------
$script:CategoryNames = @('Documents', 'Images', 'Videos', 'Audio', 'Archives', 'Executables', 'Code', 'Other', 'No extension')
$script:CategoryExtensions = [ordered]@{
    'Documents'   = 'pdf doc docx docm dot dotx odt rtf txt md markdown tex wpd pages xls xlsx xlsm xlsb xlt xltx ods csv tsv ppt pptx pptm pps ppsx odp key numbers epub mobi azw azw3 one onepkg vsd vsdx pub xps oxps msg eml log nfo'
    'Images'      = 'jpg jpeg jpe jfif png gif bmp dib tif tiff webp heic heif avif ico cur svg svgz raw cr2 cr3 nef nrw arw dng orf rw2 raf srw pef psd psb ai eps xcf jxl'
    'Videos'      = 'mp4 m4v mkv avi mov wmv flv f4v webm mpg mpeg mpe m2v m2ts mts ts vob 3gp 3g2 ogv rm rmvb asf divx'
    'Audio'       = 'mp3 wav flac aac m4a m4b ogg oga opus wma aiff aif aifc alac ape mid midi amr ac3 dts mka'
    'Archives'    = 'zip 7z rar tar gz tgz bz2 tbz2 xz txz zst lz lzma cab arj z iso img vhd vhdx vmdk wim esd dmg jar war apk'
    'Executables' = 'exe msi msix msixbundle appx appxbundle com scr dll sys ocx drv cpl bat cmd vbs vbe wsf wsh ps1 psm1 psd1 lnk reg'
    'Code'        = 'c h cpp cc cxx hpp hh cs vb fs java kt kts scala go rs py pyw rb php pl pm swift m mm js mjs cjs ts tsx jsx html htm css scss sass less json xml yaml yml toml ini cfg conf sql sh bash zsh r lua dart groovy gradle ipynb sln csproj vbproj vcxproj'
}
$script:ExtensionToCategory = New-Object 'System.Collections.Generic.Dictionary[string,int]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($categoryName in $script:CategoryExtensions.Keys) {
    $categoryIndex = [array]::IndexOf($script:CategoryNames, $categoryName)
    foreach ($ext in ($script:CategoryExtensions[$categoryName] -split '\s+')) {
        if ($ext -and -not $script:ExtensionToCategory.ContainsKey($ext)) {
            $script:ExtensionToCategory.Add($ext, $categoryIndex)
        }
    }
}
$script:OtherCategoryIndex = [array]::IndexOf($script:CategoryNames, 'Other')
$script:NoExtCategoryIndex = [array]::IndexOf($script:CategoryNames, 'No extension')

# Characters that must be escaped inside a JSON string that is embedded in an HTML
# <script> element: control chars, quote, backslash, < > & (prevents </script>
# break-out and HTML parsing tricks), U+2028/2029 and any surrogate code unit
# (Windows allows unpaired surrogates in file names; \uXXXX keeps them lossless).
$script:JsonEscapeRegex = New-Object System.Text.RegularExpressions.Regex -ArgumentList '[\x00-\x1F"\\<>&\u2028\u2029\uD800-\uDFFF]'
$script:JsonEscapeEvaluator = [System.Text.RegularExpressions.MatchEvaluator] {
    param($m)
    '\u' + ([int][char]$m.Value).ToString('x4', [System.Globalization.CultureInfo]::InvariantCulture)
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function New-UserError {
    # Creates an exception for expected, user-correctable problems (bad paths, etc.).
    # The top-level handler prints these without a stack trace.
    param([Parameter(Mandatory = $true)][string]$Message)
    $ex = New-Object System.ArgumentException -ArgumentList $Message
    $ex.Data['FolderInventoryUserError'] = $true
    return $ex
}

function Format-ByteSize {
    param([Parameter(Mandatory = $true)][double]$Bytes)
    if ($Bytes -lt 1024) { return ('{0:N0} B' -f $Bytes) }
    $units = @('KB', 'MB', 'GB', 'TB', 'PB')
    $value = $Bytes / 1024
    $i = 0
    while ($value -ge 1024 -and $i -lt $units.Count - 1) { $value = $value / 1024; $i++ }
    return ('{0:N2} {1}' -f $value, $units[$i])
}

function ConvertTo-JsonLiteral {
    param([AllowNull()][AllowEmptyString()][string]$Value)
    if ($null -eq $Value) { return 'null' }
    if ($script:JsonEscapeRegex.IsMatch($Value)) {
        return '"' + $script:JsonEscapeRegex.Replace($Value, $script:JsonEscapeEvaluator) + '"'
    }
    return '"' + $Value + '"'
}

function ConvertTo-JsonValue {
    # Small JSON serializer for metadata (strings, numbers, booleans, null,
    # dictionaries, lists). Bulk data is written by dedicated loops for speed.
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return 'null' }
    if ($Value -is [bool]) { if ($Value) { return 'true' } else { return 'false' } }
    if ($Value -is [string]) { return (ConvertTo-JsonLiteral -Value $Value) }
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [int16] -or $Value -is [byte]) { return $Value.ToString($script:Invariant) }
    if ($Value -is [double] -or $Value -is [single] -or $Value -is [decimal]) { return ([double]$Value).ToString('0.###', $script:Invariant) }
    if ($Value -is [System.Collections.IDictionary]) {
        $pairs = New-Object 'System.Collections.Generic.List[string]'
        foreach ($key in $Value.Keys) { $pairs.Add((ConvertTo-JsonLiteral -Value ([string]$key)) + ':' + (ConvertTo-JsonValue -Value $Value[$key])) }
        return '{' + ($pairs -join ',') + '}'
    }
    if ($Value -is [System.Collections.IEnumerable]) {
        $items = New-Object 'System.Collections.Generic.List[string]'
        foreach ($item in $Value) { $items.Add((ConvertTo-JsonValue -Value $item)) }
        return '[' + ($items -join ',') + ']'
    }
    return (ConvertTo-JsonLiteral -Value ([string]$Value))
}

function ConvertTo-UnixSeconds {
    param([Parameter(Mandatory = $true)][datetime]$UtcTime)
    return [long][Math]::Floor(($UtcTime.ToUniversalTime().Ticks - 621355968000000000) / 10000000)
}

function Get-CurrentFileSystemDirectory {
    # Current location of the FileSystem provider, even if the session is currently
    # located in another provider (for example HKLM:\).
    return (Get-Location -PSProvider FileSystem).ProviderPath
}

# ---------------------------------------------------------------------------
# Parameter validation
# ---------------------------------------------------------------------------
function Resolve-InventoryRoot {
    param([AllowEmptyString()][string]$Path)

    $p = $Path
    if ($null -eq $p) { $p = '' }
    # Strip whitespace and stray quotes. cmd.exe / powershell.exe -File turn  "D:\"
    # into  D:"  because \" is treated as an escaped quote.
    $p = $p.Trim().Trim('"').Trim()
    if ($p -eq '') { $p = Get-CurrentFileSystemDirectory }
    if ($p -match '^[A-Za-z]:$') { $p = $p + '\' }   # "D:" means the drive root here

    $item = $null
    try {
        $item = Get-Item -LiteralPath $p -Force -ErrorAction Stop
    } catch {
        throw (New-UserError ("RootPath not found or not accessible: '{0}'. {1}" -f $p, $_.Exception.Message))
    }
    if ($item.PSProvider.Name -ne 'FileSystem') {
        throw (New-UserError ("RootPath must be a file-system folder or drive, not '{0}' ({1} provider)." -f $p, $item.PSProvider.Name))
    }
    if (-not $item.PSIsContainer) {
        throw (New-UserError ("RootPath is a file, not a folder: '{0}'" -f $item.FullName))
    }

    $full = [System.IO.Path]::GetFullPath($item.FullName)
    $volumeRoot = [System.IO.Path]::GetPathRoot($full)
    if ($full.Length -gt $volumeRoot.Length) { $full = $full.TrimEnd('\', '/') }

    return @{
        Path  = $full
        Attr  = [int]$item.Attributes
        MTime = (ConvertTo-UnixSeconds -UtcTime $item.LastWriteTimeUtc)
    }
}

function Resolve-OutputLocation {
    param([AllowEmptyString()][string]$Path)

    $p = $Path
    if ($null -eq $p) { $p = '' }
    $p = $p.Trim().Trim('"').Trim()
    if ($p -eq '') { $p = Join-Path -Path (Get-CurrentFileSystemDirectory) -ChildPath 'FolderInventory.html' }

    try {
        $full = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($p)
    } catch {
        throw (New-UserError ("OutputHtml is not a valid path: '{0}'. {1}" -f $p, $_.Exception.Message))
    }
    $full = [System.IO.Path]::GetFullPath($full)

    if ([System.IO.Directory]::Exists($full)) {
        Write-Host ("OutputHtml is a folder; the report will be written as FolderInventory.html inside it.") -ForegroundColor Yellow
        $full = [System.IO.Path]::Combine($full, 'FolderInventory.html')
    }
    $ext = [System.IO.Path]::GetExtension($full)
    if ($ext -ne '.html' -and $ext -ne '.htm') {
        Write-Warning ("OutputHtml does not end in .html: '{0}'. The file will still be HTML." -f $full)
    }

    $parent = [System.IO.Path]::GetDirectoryName($full)
    if (-not $parent) { throw (New-UserError ("OutputHtml has no parent folder: '{0}'" -f $full)) }
    if (-not [System.IO.Directory]::Exists($parent)) {
        try {
            [void][System.IO.Directory]::CreateDirectory($parent)
            Write-Host ("Created output folder: {0}" -f $parent)
        } catch {
            throw (New-UserError ("Cannot create output folder '{0}': {1}" -f $parent, $_.Exception.Message))
        }
    }

    # Fail fast (before a long scan) if the report cannot be written.
    if ([System.IO.File]::Exists($full)) {
        $existing = New-Object System.IO.FileInfo -ArgumentList $full
        if ($existing.IsReadOnly) { throw (New-UserError ("OutputHtml exists and is read-only: '{0}'" -f $full)) }
    }
    $probe = [System.IO.Path]::Combine($parent, ('.FolderInventory-write-test-' + [guid]::NewGuid().ToString('N') + '.tmp'))
    try {
        [System.IO.File]::WriteAllText($probe, 'test')
        [System.IO.File]::Delete($probe)
    } catch {
        throw (New-UserError ("Cannot write to output folder '{0}': {1}" -f $parent, $_.Exception.Message))
    }

    return $full
}

# ---------------------------------------------------------------------------
# Scanning (runs in a bounded runspace pool)
# ---------------------------------------------------------------------------
# Work item:   @(displayPath, parentDisplayPath, depth, attributes, mtimeUnixSeconds)
# Result:      hashtable, Kind = 'Dir' (one per folder) or 'WorkerFatal'
# File record: @(name, sizeBytes, mtimeUnix, ctimeUnix, attributes, extension)
$script:WorkerScript = {
    param(
        $WorkQueue,
        $ResultQueue,
        $Pending,
        $InFlight,
        $Control,
        [bool]$IncludeHidden,
        [bool]$IncludeSystem,
        [bool]$FollowReparsePoints,
        [int]$MaxDepth,
        [int]$WorkerId
    )
    $ErrorActionPreference = 'Stop'
    Set-StrictMode -Off
    $epochTicks = 621355968000000000
    $isWinFs = ([System.IO.Path]::DirectorySeparatorChar -eq '\')

    function ConvertTo-DisplayPath([string]$p) {
        if ($p.StartsWith('\\?\UNC\')) { return '\\' + $p.Substring(8) }
        if ($p.StartsWith('\\?\')) { return $p.Substring(4) }
        return $p
    }
    function ConvertTo-ExtendedPath([string]$p) {
        if (-not $isWinFs -or $p.StartsWith('\\?\')) { return $p }
        if ($p.StartsWith('\\')) { return '\\?\UNC\' + $p.Substring(2) }
        return '\\?\' + $p
    }
    function Get-FsErrorInfo($err) {
        # Unwrap PowerShell's MethodInvocationException etc. and classify.
        $ex = $err
        if ($ex -is [System.Management.Automation.ErrorRecord]) { $ex = $ex.Exception }
        while ($null -ne $ex.InnerException -and $ex -is [System.Management.Automation.RuntimeException]) { $ex = $ex.InnerException }
        $type = 'Other'
        if ($ex -is [System.UnauthorizedAccessException] -or $ex -is [System.Security.SecurityException]) { $type = 'AccessDenied' }
        elseif ($ex -is [System.IO.DirectoryNotFoundException] -or $ex -is [System.IO.FileNotFoundException]) { $type = 'NotFound' }
        elseif ($ex -is [System.IO.PathTooLongException]) { $type = 'PathTooLong' }
        elseif ($ex -is [System.IO.IOException]) { $type = 'IOError' }
        $msg = [string]$ex.Message
        if ($type -eq 'Other' -or $type -eq 'IOError') { $msg = $ex.GetType().FullName + ': ' + $msg }
        return , @($type, $msg)
    }

    try {
        while (-not $Pending.IsSet -and -not $Control.Cancel) {
            $item = $null
            if (-not $WorkQueue.TryDequeue([ref]$item)) {
                Start-Sleep -Milliseconds 10
                continue
            }
            $displayPath = [string]$item[0]
            [void]$InFlight.TryAdd($displayPath, $item)
            $resultPushed = $false
            try {
                $depth = [int]$item[2]
                $res = @{
                    Kind = 'Dir'; Path = $displayPath; Parent = $item[1]; Depth = $depth
                    Attr = $item[3]; MTime = $item[4]; Status = 'OK'; ErrType = $null; ErrMsg = $null
                    Files = $null; FileErrors = $null; Reparse = $null; ExclFiles = 0; ExclDirs = 0
                }
                $children = $null

                if ($depth -gt $MaxDepth) {
                    $res.Status = 'Skipped'; $res.ErrType = 'MaxDepth'
                    $res.ErrMsg = ('Folder nesting deeper than {0} levels; not scanned.' -f $MaxDepth)
                } else {
                    $files = New-Object 'System.Collections.Generic.List[object]'
                    $children = New-Object 'System.Collections.Generic.List[object]'
                    $scanPath = $displayPath
                    if ($isWinFs -and $scanPath.Length -ge 240) { $scanPath = ConvertTo-ExtendedPath $scanPath }

                    while ($true) {
                        $readAny = $false
                        try {
                            $di = New-Object System.IO.DirectoryInfo -ArgumentList $scanPath
                            foreach ($e in $di.EnumerateFileSystemInfos()) {
                                $readAny = $true
                                try {
                                    $a = [int]$e.Attributes
                                    if (((($a -band 2) -ne 0) -and -not $IncludeHidden) -or ((($a -band 4) -ne 0) -and -not $IncludeSystem)) {
                                        if (($a -band 16) -ne 0) { $res.ExclDirs++ } else { $res.ExclFiles++ }
                                        continue
                                    }
                                    $mt = [long][Math]::Floor(($e.LastWriteTimeUtc.Ticks - $epochTicks) / 10000000)
                                    if (($a -band 16) -ne 0) {
                                        if ((($a -band 1024) -ne 0) -and -not $FollowReparsePoints) {
                                            if ($null -eq $res.Reparse) { $res.Reparse = New-Object 'System.Collections.Generic.List[string]' }
                                            $res.Reparse.Add((ConvertTo-DisplayPath $e.FullName))
                                            continue
                                        }
                                        $children.Add(@($e.FullName, $a, $mt))
                                    } else {
                                        $n = $e.Name
                                        $dot = $n.LastIndexOf('.')
                                        $ext = ''
                                        if ($dot -ge 0 -and $dot -lt ($n.Length - 1)) { $ext = [string]::Intern($n.Substring($dot + 1).ToLowerInvariant()) }
                                        $ct = [long][Math]::Floor(($e.CreationTimeUtc.Ticks - $epochTicks) / 10000000)
                                        $files.Add(@($n, [long]$e.Length, $mt, $ct, $a, $ext))
                                    }
                                } catch {
                                    $fi = Get-FsErrorInfo $_
                                    $ep = $displayPath
                                    try { $ep = ConvertTo-DisplayPath $e.FullName } catch { $ep = $displayPath + ' (entry name unavailable)' }
                                    if ($null -eq $res.FileErrors) { $res.FileErrors = New-Object 'System.Collections.Generic.List[object]' }
                                    $res.FileErrors.Add(@($ep, $fi[0], $fi[1]))
                                }
                            }
                            break
                        } catch {
                            $info = Get-FsErrorInfo $_
                            if (-not $readAny -and $info[0] -eq 'PathTooLong' -and $isWinFs -and -not $scanPath.StartsWith('\\?\')) {
                                $scanPath = ConvertTo-ExtendedPath $scanPath
                                continue
                            }
                            if ($readAny) { $res.Status = 'Partial' } else { $res.Status = 'Skipped' }
                            $res.ErrType = $info[0]
                            $res.ErrMsg = $info[1]
                            break
                        }
                    }
                    $res.Files = $files
                }

                $ResultQueue.Enqueue($res)
                $resultPushed = $true

                if ($null -ne $children) {
                    foreach ($c in $children) {
                        $childItem = @((ConvertTo-DisplayPath $c[0]), $displayPath, ($depth + 1), $c[1], $c[2])
                        $Pending.AddCount(1)
                        $WorkQueue.Enqueue($childItem)
                    }
                }
            } catch {
                $info = Get-FsErrorInfo $_
                $line = 0
                try { $line = $_.InvocationInfo.ScriptLineNumber } catch { $line = 0 }
                if (-not $resultPushed) {
                    $ResultQueue.Enqueue(@{
                            Kind = 'Dir'; Path = $displayPath; Parent = $item[1]; Depth = [int]$item[2]
                            Attr = $item[3]; MTime = $item[4]; Status = 'Skipped'; ErrType = 'WorkerError'
                            ErrMsg = ('Unexpected error in worker {0} (worker line {1}): {2}' -f $WorkerId, $line, $info[1])
                            Files = $null; FileErrors = $null; Reparse = $null; ExclFiles = 0; ExclDirs = 0
                        })
                } else {
                    $ResultQueue.Enqueue(@{ Kind = 'WorkerFatal'; Worker = $WorkerId; Path = $displayPath
                            Message = ('Error after listing folder (worker line {0}); some subfolders may not have been queued: {1}' -f $line, $info[1]) })
                }
            } finally {
                $removed = $null
                [void]$InFlight.TryRemove($displayPath, [ref]$removed)
                [void]$Pending.Signal()
            }
        }
    } catch {
        $msg = $_.Exception.Message
        $ResultQueue.Enqueue(@{ Kind = 'WorkerFatal'; Worker = $WorkerId; Path = $null; Message = ('Worker {0} stopped unexpectedly: {1}' -f $WorkerId, $msg) })
    }
}

function New-DirectoryRecord {
    # Record for a folder that was discovered but never listed (timeout, worker stop).
    param($Item, [string]$Status, [string]$ErrType, [string]$ErrMsg)
    return @{
        Kind = 'Dir'; Path = [string]$Item[0]; Parent = $Item[1]; Depth = [int]$Item[2]
        Attr = $Item[3]; MTime = $Item[4]; Status = $Status; ErrType = $ErrType; ErrMsg = $ErrMsg
        Files = $null; FileErrors = $null; Reparse = $null; ExclFiles = 0; ExclDirs = 0
    }
}

function Receive-ScanResult {
    # Moves finished folder results from the shared queue into the main-thread store.
    param([Parameter(Mandatory = $true)][hashtable]$Context)
    $r = $null
    while ($Context.ResultQueue.TryDequeue([ref]$r)) {
        if ($r.Kind -ne 'Dir') {
            $Context.WorkerErrors.Add($r)
            continue
        }
        if ($Context.Index.ContainsKey($r.Path)) {
            # Should never happen (each folder is queued once); never double count.
            $Context.Internal.Add(@($r.Path, 'Internal', 'Duplicate folder result received and ignored; files were counted once.'))
            continue
        }
        $Context.Index.Add($r.Path, $Context.Dirs.Count)
        $Context.Dirs.Add($r)
        if ($null -ne $r.Files) { $Context.FileCount += $r.Files.Count }
        if ($r.Status -ne 'OK' -or $null -ne $r.FileErrors) { $Context.ProblemCount++ }
    }
}

function Invoke-FolderScan {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Root,
        [Parameter(Mandatory = $true)][int]$Workers,
        [Parameter(Mandatory = $true)][bool]$Hidden,
        [Parameter(Mandatory = $true)][bool]$System,
        [Parameter(Mandatory = $true)][bool]$FollowLinks,
        [Parameter(Mandatory = $true)][int]$TimeLimitMinutes
    )

    if ($script:IsWindowsFileSystem) { $comparer = [System.StringComparer]::OrdinalIgnoreCase } else { $comparer = [System.StringComparer]::Ordinal }
    $ctx = @{
        WorkQueue    = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'
        ResultQueue  = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'
        InFlight     = New-Object 'System.Collections.Concurrent.ConcurrentDictionary[string,object]' -ArgumentList $comparer
        Pending      = New-Object System.Threading.CountdownEvent -ArgumentList 1
        Control      = [hashtable]::Synchronized(@{ Cancel = $false })
        Dirs         = New-Object 'System.Collections.Generic.List[object]'
        Index        = New-Object 'System.Collections.Generic.Dictionary[string,int]' -ArgumentList $comparer
        WorkerErrors = New-Object 'System.Collections.Generic.List[object]'
        Internal     = New-Object 'System.Collections.Generic.List[object]'
        FileCount    = [long]0
        ProblemCount = 0
        TimedOut     = $false
        Comparer     = $comparer
    }
    $ctx.WorkQueue.Enqueue(@($Root.Path, $null, 0, $Root.Attr, $Root.MTime))

    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $pool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $Workers, $iss, $Host)
    $pool.Open()
    $workerList = New-Object 'System.Collections.Generic.List[object]'
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $activity = 'Scanning ' + $Root.Path
    try {
        for ($w = 1; $w -le $Workers; $w++) {
            $ps = [System.Management.Automation.PowerShell]::Create()
            $ps.RunspacePool = $pool
            [void]$ps.AddScript($script:WorkerScript.ToString())
            [void]$ps.AddParameter('WorkQueue', $ctx.WorkQueue)
            [void]$ps.AddParameter('ResultQueue', $ctx.ResultQueue)
            [void]$ps.AddParameter('Pending', $ctx.Pending)
            [void]$ps.AddParameter('InFlight', $ctx.InFlight)
            [void]$ps.AddParameter('Control', $ctx.Control)
            [void]$ps.AddParameter('IncludeHidden', $Hidden)
            [void]$ps.AddParameter('IncludeSystem', $System)
            [void]$ps.AddParameter('FollowReparsePoints', $FollowLinks)
            [void]$ps.AddParameter('MaxDepth', $script:MaxDepth)
            [void]$ps.AddParameter('WorkerId', $w)
            $handle = $ps.BeginInvoke()
            $workerList.Add(@{ Id = $w; PS = $ps; Handle = $handle })
        }

        $lastProgress = [DateTime]::MinValue
        $lastLog = [DateTime]::UtcNow
        $cancelRequestedAt = $null
        $stoppedByForce = $false
        while ($true) {
            Receive-ScanResult -Context $ctx
            $running = 0
            foreach ($wk in $workerList) { if (-not $wk.Handle.IsCompleted) { $running++ } }
            if ($running -eq 0) { break }

            if ($TimeLimitMinutes -gt 0 -and -not $ctx.TimedOut -and $stopwatch.Elapsed.TotalMinutes -ge $TimeLimitMinutes) {
                $ctx.TimedOut = $true
                $ctx.Control.Cancel = $true
                $cancelRequestedAt = [DateTime]::UtcNow
                Write-Warning ("Time limit of {0} minute(s) reached; stopping the scan and reporting partial results." -f $TimeLimitMinutes)
            }
            if ($null -ne $cancelRequestedAt -and -not $stoppedByForce -and ([DateTime]::UtcNow - $cancelRequestedAt).TotalSeconds -ge 60) {
                # Workers stuck in a slow folder (e.g. an unresponsive network share).
                $stoppedByForce = $true
                foreach ($wk in $workerList) {
                    if (-not $wk.Handle.IsCompleted) {
                        $ctx.WorkerErrors.Add(@{ Kind = 'WorkerFatal'; Worker = $wk.Id; Path = $null; Message = ('Worker {0} did not stop within 60 seconds after the time limit and was terminated.' -f $wk.Id) })
                        try { $wk.PS.Stop() } catch { Write-Verbose ("Stopping worker {0}: {1}" -f $wk.Id, $_.Exception.Message) }
                    }
                }
            }

            $now = [DateTime]::UtcNow
            if (($now - $lastProgress).TotalMilliseconds -ge 1000) {
                $lastProgress = $now
                $status = ('Folders: {0:N0} | Files: {1:N0} | Queued: {2:N0} | Problems: {3:N0} | Elapsed: {4:hh\:mm\:ss}' -f $ctx.Dirs.Count, $ctx.FileCount, $ctx.WorkQueue.Count, $ctx.ProblemCount, $stopwatch.Elapsed)
                Write-Progress -Activity $activity -Status $status
                if (($now - $lastLog).TotalSeconds -ge 30) {
                    $lastLog = $now
                    Write-Host ('  ... ' + $status)
                }
            }
            Start-Sleep -Milliseconds 200
        }
        Receive-ScanResult -Context $ctx

        foreach ($wk in $workerList) {
            try {
                [void]$wk.PS.EndInvoke($wk.Handle)
            } catch {
                $ctx.WorkerErrors.Add(@{ Kind = 'WorkerFatal'; Worker = $wk.Id; Path = $null; Message = ('Worker {0} ended with an error: {1}' -f $wk.Id, $_.Exception.Message) })
            }
            if ($wk.PS.Streams.Error.Count -gt 0) {
                foreach ($er in $wk.PS.Streams.Error) {
                    $ctx.WorkerErrors.Add(@{ Kind = 'WorkerFatal'; Worker = $wk.Id; Path = $null; Message = ('Worker {0} reported: {1}' -f $wk.Id, $er.ToString()) })
                }
            }
        }
        Receive-ScanResult -Context $ctx

        # Folders that were discovered but never listed.
        if ($ctx.TimedOut) { $reasonType = 'Timeout'; $reasonMsg = 'Not scanned: the time limit was reached.' }
        else { $reasonType = 'NotScanned'; $reasonMsg = 'Not scanned: the scanning worker stopped before reaching this folder.' }
        foreach ($key in @($ctx.InFlight.Keys)) {
            $it = $null
            if ($ctx.InFlight.TryRemove($key, [ref]$it) -and -not $ctx.Index.ContainsKey([string]$it[0])) {
                $rec = New-DirectoryRecord -Item $it -Status 'NotScanned' -ErrType $reasonType -ErrMsg ('Listing was interrupted. ' + $reasonMsg)
                $ctx.Index.Add($rec.Path, $ctx.Dirs.Count); $ctx.Dirs.Add($rec); $ctx.ProblemCount++
            }
        }
        $left = $null
        while ($ctx.WorkQueue.TryDequeue([ref]$left)) {
            if (-not $ctx.Index.ContainsKey([string]$left[0])) {
                $rec = New-DirectoryRecord -Item $left -Status 'NotScanned' -ErrType $reasonType -ErrMsg $reasonMsg
                $ctx.Index.Add($rec.Path, $ctx.Dirs.Count); $ctx.Dirs.Add($rec); $ctx.ProblemCount++
            }
        }
    } finally {
        Write-Progress -Activity $activity -Completed
        foreach ($wk in $workerList) {
            if (-not $wk.Handle.IsCompleted) {
                $ctx.Control.Cancel = $true
                try { $wk.PS.Stop() } catch { Write-Verbose ("Stopping worker {0}: {1}" -f $wk.Id, $_.Exception.Message) }
            }
            $wk.PS.Dispose()
        }
        $pool.Close()
        $pool.Dispose()
        $ctx.Pending.Dispose()
    }
    $ctx.ElapsedSeconds = $stopwatch.Elapsed.TotalSeconds
    return $ctx
}

# ---------------------------------------------------------------------------
# Aggregation (single bottom-up pass, no rescanning)
# ---------------------------------------------------------------------------
function Build-InventoryModel {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Scan,
        [Parameter(Mandatory = $true)][hashtable]$Root
    )
    $dirs = $Scan.Dirs
    $n = $dirs.Count
    if ($n -eq 0) { throw 'No scan results were produced for the root folder.' }
    if ($Scan.Comparer.Compare([string]$dirs[0].Path, $Root.Path) -ne 0) {
        throw ("Internal error: first result '{0}' is not the root '{1}'." -f $dirs[0].Path, $Root.Path)
    }
    $rootRec = $dirs[0]
    if ($rootRec.Status -eq 'Skipped' -or $rootRec.Status -eq 'NotScanned') {
        throw (New-UserError ("Cannot read the root folder '{0}': [{1}] {2}" -f $Root.Path, $rootRec.ErrType, $rootRec.ErrMsg))
    }

    $parent = New-Object int[] $n
    $directSize = New-Object long[] $n
    $directFiles = New-Object long[] $n
    $totalSize = New-Object long[] $n
    $totalFiles = New-Object long[] $n
    $subdirs = New-Object int[] $n
    $descDirs = New-Object int[] $n
    $partial = New-Object bool[] $n
    $depthKey = New-Object int[] $n
    $order = New-Object int[] $n

    $extIndex = New-Object 'System.Collections.Generic.Dictionary[string,int]' ([System.StringComparer]::Ordinal)
    $extNames = New-Object 'System.Collections.Generic.List[string]'
    $extCount = New-Object 'System.Collections.Generic.List[long]'
    $extSize = New-Object 'System.Collections.Generic.List[long]'

    $largestSize = [long]-1; $largestDir = -1; $largestName = $null
    $sumFileSize = [long]0; $sumFileCount = [long]0

    for ($i = 0; $i -lt $n; $i++) {
        $d = $dirs[$i]
        if ($i -eq 0) {
            $parent[$i] = -1
        } else {
            $p = 0
            if ($null -ne $d.Parent -and $Scan.Index.TryGetValue([string]$d.Parent, [ref]$p)) {
                $parent[$i] = $p
            } else {
                $parent[$i] = 0
                $Scan.Internal.Add(@($d.Path, 'Internal', ('Parent folder record not found ({0}); attached to root.' -f $d.Parent)))
            }
        }
        $depthKey[$i] = - [int]$d.Depth
        $order[$i] = $i
        $partial[$i] = ($d.Status -ne 'OK') -or ($null -ne $d.FileErrors -and $d.FileErrors.Count -gt 0)

        if ($null -ne $d.Files) {
            $ds = [long]0
            foreach ($f in $d.Files) {
                $size = [long]$f[1]
                $ds += $size
                $ext = [string]$f[5]
                $ei = 0
                if (-not $extIndex.TryGetValue($ext, [ref]$ei)) {
                    $ei = $extNames.Count
                    $extIndex.Add($ext, $ei); $extNames.Add($ext); $extCount.Add(0); $extSize.Add(0)
                }
                $f[5] = $ei
                $extCount[$ei] = $extCount[$ei] + 1
                $extSize[$ei] = $extSize[$ei] + $size
                if ($size -gt $largestSize) { $largestSize = $size; $largestDir = $i; $largestName = [string]$f[0] }
            }
            $directSize[$i] = $ds
            $directFiles[$i] = $d.Files.Count
            $sumFileSize += $ds
            $sumFileCount += $d.Files.Count
        }
        $totalSize[$i] = $directSize[$i]
        $totalFiles[$i] = $directFiles[$i]
    }

    # Deepest folders first, so each child is complete before it is added to its parent.
    [System.Array]::Sort($depthKey, $order)
    foreach ($i in $order) {
        $p = $parent[$i]
        if ($p -lt 0) { continue }
        $totalSize[$p] += $totalSize[$i]
        $totalFiles[$p] += $totalFiles[$i]
        $subdirs[$p]++
        $descDirs[$p] += $descDirs[$i] + 1
        if ($partial[$i]) { $partial[$p] = $true }
    }

    # Consistency checks: the totals must equal the sum of the collected file records.
    if ($totalSize[0] -ne $sumFileSize -or $totalFiles[0] -ne $sumFileCount -or $sumFileCount -ne $Scan.FileCount) {
        $Scan.Internal.Add(@($Root.Path, 'Internal', ('Consistency check failed: root total {0} bytes/{1} files, records {2} bytes/{3} files, received {4} files.' -f $totalSize[0], $totalFiles[0], $sumFileSize, $sumFileCount, $Scan.FileCount)))
    }

    # Extension table: [ext, categoryIndex]
    $extCategory = New-Object int[] $extNames.Count
    for ($e = 0; $e -lt $extNames.Count; $e++) {
        $name = $extNames[$e]
        if ($name -eq '') { $extCategory[$e] = $script:NoExtCategoryIndex }
        else {
            $ci = 0
            if ($script:ExtensionToCategory.TryGetValue($name, [ref]$ci)) { $extCategory[$e] = $ci } else { $extCategory[$e] = $script:OtherCategoryIndex }
        }
    }

    # Problems list: @(path, type, message, kind)
    $problems = New-Object 'System.Collections.Generic.List[object]'
    $reparse = New-Object 'System.Collections.Generic.List[string]'
    $skippedFolders = 0; $partialFolders = 0; $fileErrors = 0; $exclFiles = [long]0; $exclDirs = [long]0
    foreach ($d in $dirs) {
        switch ($d.Status) {
            'Skipped' { $skippedFolders++; $problems.Add(@($d.Path, [string]$d.ErrType, [string]$d.ErrMsg, 'Folder skipped')) }
            'NotScanned' { $skippedFolders++; $problems.Add(@($d.Path, [string]$d.ErrType, [string]$d.ErrMsg, 'Folder not scanned')) }
            'Partial' { $partialFolders++; $problems.Add(@($d.Path, [string]$d.ErrType, [string]$d.ErrMsg, 'Folder partially listed')) }
        }
        if ($null -ne $d.FileErrors) {
            foreach ($fe in $d.FileErrors) { $fileErrors++; $problems.Add(@($fe[0], $fe[1], $fe[2], 'File')) }
        }
        if ($null -ne $d.Reparse) { foreach ($rp in $d.Reparse) { $reparse.Add($rp) } }
        $exclFiles += $d.ExclFiles
        $exclDirs += $d.ExclDirs
    }
    foreach ($we in $Scan.WorkerErrors) { $problems.Add(@([string]$we.Path, 'WorkerError', [string]$we.Message, 'Worker')) }
    foreach ($ie in $Scan.Internal) { $problems.Add(@([string]$ie[0], [string]$ie[1], [string]$ie[2], 'Internal')) }

    $complete = ($problems.Count -eq 0) -and (-not $Scan.TimedOut)

    $largest = $null
    if ($largestDir -ge 0) {
        $largest = @{ Name = $largestName; Folder = $dirs[$largestDir].Path; Size = $largestSize }
    }
    $topExt = $null; $topExtCount = [long]0
    for ($e = 0; $e -lt $extNames.Count; $e++) { if ($extCount[$e] -gt $topExtCount) { $topExtCount = $extCount[$e]; $topExt = $extNames[$e] } }

    return @{
        Dirs = $dirs; Parent = $parent; DirectSize = $directSize; DirectFiles = $directFiles
        TotalSize = $totalSize; TotalFiles = $totalFiles; Subdirs = $subdirs; DescDirs = $descDirs; Partial = $partial
        ExtNames = $extNames; ExtCategory = $extCategory; ExtCount = $extCount; ExtSize = $extSize
        Problems = $problems; Reparse = $reparse
        SkippedFolders = $skippedFolders; PartialFolders = $partialFolders; FileErrors = $fileErrors
        WorkerErrors = $Scan.WorkerErrors.Count; InternalErrors = $Scan.Internal.Count
        ExcludedFiles = $exclFiles; ExcludedDirs = $exclDirs
        Complete = $complete; TimedOut = $Scan.TimedOut
        FileCount = $sumFileCount; TotalBytes = $sumFileSize; FolderCount = $n
        Largest = $largest; TopExtension = $topExt; TopExtensionCount = $topExtCount
        ElapsedSeconds = $Scan.ElapsedSeconds; RootPath = $Root.Path
    }
}

# ---------------------------------------------------------------------------
# HTML report generation (separate from scanning)
# ---------------------------------------------------------------------------
$script:HtmlTemplate = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; img-src data:; connect-src 'none'; form-action 'none'; base-uri 'none'">
<title>Folder Inventory</title>
<style>
:root {
  --bg: #f6f7f9; --panel: #ffffff; --text: #1d232b; --muted: #5d6875; --line: #e1e5ea; --line2: #eef1f4;
  --accent: #2f6fdb; --accent-soft: #e6eefc; --ok: #1d7f4e; --ok-soft: #e3f4ea; --warn: #9a5b00; --warn-soft: #fdf1dc;
  --bad: #b3261e; --bad-soft: #fbe5e3; --bar: #7ea6ec; --row-hover: #f3f6fb; --sel: #dfe9fb;
  --mono: Consolas, "Cascadia Mono", "SFMono-Regular", Menlo, monospace;
  --sans: "Segoe UI", system-ui, -apple-system, Roboto, "Helvetica Neue", Arial, sans-serif;
}
@media (prefers-color-scheme: dark) {
  :root {
    --bg: #15181c; --panel: #1d2126; --text: #e4e7eb; --muted: #9aa4af; --line: #313840; --line2: #272c33;
    --accent: #6ea0ff; --accent-soft: #23314a; --ok: #5cc98d; --ok-soft: #1c3428; --warn: #f0b45a; --warn-soft: #3a2e1a;
    --bad: #ff8a80; --bad-soft: #3d2220; --bar: #4f79c4; --row-hover: #242a31; --sel: #263650;
  }
}
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; }
body { background: var(--bg); color: var(--text); font: 14px/1.45 var(--sans); }
button, input, select { font: inherit; color: inherit; }
header.top { padding: 18px 24px 10px; }
header.top h1 { margin: 0 0 4px; font-size: 20px; font-weight: 600; display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
.meta { color: var(--muted); font-size: 13px; display: flex; flex-wrap: wrap; gap: 4px 18px; }
.meta b { color: var(--text); font-weight: 600; word-break: break-all; }
.badge { display: inline-block; font-size: 12px; font-weight: 600; padding: 2px 8px; border-radius: 10px; line-height: 1.5; white-space: nowrap; }
.badge.ok { background: var(--ok-soft); color: var(--ok); }
.badge.warn { background: var(--warn-soft); color: var(--warn); }
.badge.bad { background: var(--bad-soft); color: var(--bad); }
.badge.info { background: var(--accent-soft); color: var(--accent); }
.badge.mute { background: var(--line2); color: var(--muted); }
.banner { margin: 0 24px 12px; padding: 10px 14px; border-radius: 8px; background: var(--warn-soft); color: var(--warn); }
.banner.bad { background: var(--bad-soft); color: var(--bad); }
.cards { display: grid; grid-template-columns: repeat(auto-fill, minmax(155px, 1fr)); gap: 12px; padding: 0 24px 14px; }
.card { background: var(--panel); border: 1px solid var(--line); border-radius: 10px; padding: 12px 14px; min-width: 0; }
.card .k { color: var(--muted); font-size: 12px; text-transform: uppercase; letter-spacing: .04em; }
.card .v { font-size: 20px; font-weight: 600; margin-top: 2px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.card .s { color: var(--muted); font-size: 12px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
nav.tabs { display: flex; gap: 4px; padding: 0 24px; border-bottom: 1px solid var(--line); }
nav.tabs button { background: none; border: 0; border-bottom: 2px solid transparent; padding: 8px 14px; cursor: pointer; color: var(--muted); font-weight: 600; }
nav.tabs button[aria-selected="true"] { color: var(--accent); border-bottom-color: var(--accent); }
main { padding: 14px 24px 24px; }
.tab[hidden] { display: none !important; }
.layout { display: grid; grid-template-columns: minmax(280px, 30%) 1fr; gap: 14px; align-items: start; }
@media (max-width: 1000px) { .layout { grid-template-columns: 1fr; } }
.panel { background: var(--panel); border: 1px solid var(--line); border-radius: 10px; min-width: 0; }
.panel > .hd { padding: 10px 12px; border-bottom: 1px solid var(--line); display: flex; gap: 6px; align-items: center; flex-wrap: wrap; }
.panel > .hd h2 { margin: 0 auto 0 0; font-size: 14px; }
.panel > .bd { padding: 10px 12px; }
.btn { background: var(--panel); border: 1px solid var(--line); border-radius: 6px; padding: 4px 10px; cursor: pointer; white-space: nowrap; }
.btn:hover { border-color: var(--accent); }
.btn.primary { background: var(--accent); border-color: var(--accent); color: #fff; }
.btn:disabled { opacity: .5; cursor: default; }
.link { background: none; border: 0; padding: 0; color: var(--accent); cursor: pointer; text-align: left; }
.link:hover { text-decoration: underline; }
input[type="text"], input[type="search"], input[type="number"], input[type="date"], select {
  background: var(--panel); border: 1px solid var(--line); border-radius: 6px; padding: 4px 8px; min-width: 0;
}
#tree { max-height: 72vh; overflow: auto; font-size: 13px; }
.trow { display: flex; align-items: center; gap: 4px; padding: 2px 4px; border-radius: 4px; white-space: nowrap; }
.trow:hover { background: var(--row-hover); }
.trow.sel { background: var(--sel); }
.ttog { width: 18px; height: 18px; border: 0; background: none; cursor: pointer; padding: 0; color: var(--muted); flex: none; line-height: 18px; }
.tspace { width: 18px; flex: none; }
.tname { background: none; border: 0; padding: 0; cursor: pointer; overflow: hidden; text-overflow: ellipsis; min-width: 40px; text-align: left; }
.tname:hover { color: var(--accent); text-decoration: underline; }
.tstat { color: var(--muted); font-size: 12px; margin-left: auto; padding-left: 8px; font-variant-numeric: tabular-nums; }
.tmore { padding: 2px 4px 2px 0; }
.filters { display: grid; grid-template-columns: 1fr; gap: 10px; }
.frow { display: flex; flex-wrap: wrap; gap: 8px; align-items: center; }
.frow > label.lbl { color: var(--muted); font-size: 12px; min-width: 86px; }
.chip { border: 1px solid var(--line); background: var(--panel); border-radius: 14px; padding: 2px 10px; cursor: pointer; font-size: 13px; white-space: nowrap; }
.chip[aria-pressed="true"] { background: var(--accent); border-color: var(--accent); color: #fff; }
.chip .n { opacity: .7; font-size: 11px; margin-left: 4px; }
details.extpick { position: relative; }
details.extpick > summary { list-style: none; cursor: pointer; }
details.extpick > summary::-webkit-details-marker { display: none; }
.extbox { position: absolute; z-index: 20; top: 32px; left: 0; width: 320px; max-width: 80vw; background: var(--panel); border: 1px solid var(--line); border-radius: 8px; box-shadow: 0 6px 24px rgba(0,0,0,.18); padding: 8px; }
.extlist { max-height: 300px; overflow: auto; margin-top: 6px; }
.extlist label { display: flex; gap: 6px; align-items: center; padding: 2px 4px; cursor: pointer; border-radius: 4px; }
.extlist label:hover { background: var(--row-hover); }
.extlist .n { margin-left: auto; color: var(--muted); font-size: 12px; font-variant-numeric: tabular-nums; }
#search { width: min(420px, 100%); }
#minVal { width: 90px; }
#activeFilters { display: flex; flex-wrap: wrap; gap: 6px; }
.af { background: var(--accent-soft); color: var(--accent); border-radius: 12px; padding: 2px 4px 2px 10px; font-size: 12px; display: inline-flex; align-items: center; gap: 4px; max-width: 100%; }
.af span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 480px; }
.af button { border: 0; background: none; color: inherit; cursor: pointer; font-size: 14px; line-height: 1; padding: 0 4px; }
.tablebar { display: flex; flex-wrap: wrap; gap: 8px 14px; align-items: center; padding: 8px 12px; border-top: 1px solid var(--line); border-bottom: 1px solid var(--line); }
.tablebar .count { margin-right: auto; }
.pager { display: flex; gap: 4px; align-items: center; }
.pager input { width: 64px; }
.tablewrap { overflow-x: auto; }
table { border-collapse: collapse; width: 100%; }
th, td { text-align: left; padding: 5px 10px; border-bottom: 1px solid var(--line2); vertical-align: top; }
th { font-size: 12px; color: var(--muted); font-weight: 600; white-space: nowrap; user-select: none; position: sticky; top: 0; background: var(--panel); }
th.sortable { cursor: pointer; }
th.sortable:hover { color: var(--accent); }
th .arrow { display: inline-block; width: 12px; }
td.num, th.num { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
td.date { white-space: nowrap; font-variant-numeric: tabular-nums; }
td.name { font-weight: 500; word-break: break-word; min-width: 140px; }
td.path { word-break: break-all; min-width: 160px; }
td.path .folder { color: var(--muted); }
td.full { font-family: var(--mono); font-size: 12px; word-break: break-all; }
td.type .cat { color: var(--muted); font-size: 12px; }
td.attr { font-family: var(--mono); font-size: 12px; color: var(--muted); }
tbody tr:hover { background: var(--row-hover); }
.empty { padding: 28px; text-align: center; color: var(--muted); }
.grid2 { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }
@media (max-width: 1100px) { .grid2 { grid-template-columns: 1fr; } }
.bar { height: 8px; background: var(--line2); border-radius: 4px; overflow: hidden; min-width: 60px; }
.bar > i { display: block; height: 100%; background: var(--bar); }
.note { color: var(--muted); font-size: 12px; }
footer { padding: 4px 24px 24px; color: var(--muted); font-size: 12px; }
.mono { font-family: var(--mono); font-size: 12px; word-break: break-all; }
.toolbar { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; margin-bottom: 8px; }
#fatal { margin: 24px; padding: 16px; border-radius: 8px; background: var(--bad-soft); color: var(--bad); }
</style>
</head>
<body>
<div id="fatal" hidden></div>
<noscript><div style="margin:24px;padding:16px;background:#fde">This report needs JavaScript enabled to display the inventory.</div></noscript>
<div id="app" hidden>
  <header class="top">
    <h1><span>Folder Inventory</span> <span id="statusBadge" class="badge"></span></h1>
    <div class="meta">
      <span>Root: <b id="mRoot"></b></span>
      <span>Generated: <b id="mGen"></b></span>
      <span>Scan duration: <b id="mDur"></b></span>
      <span>Options: <b id="mOpt"></b></span>
    </div>
  </header>
  <div id="banner" class="banner" hidden></div>
  <section class="cards" id="cards" aria-label="Summary"></section>
  <nav class="tabs" role="tablist">
    <button role="tab" data-tab="files" aria-selected="true">Files</button>
    <button role="tab" data-tab="analysis" aria-selected="false">Size analysis</button>
    <button role="tab" data-tab="problems" aria-selected="false" id="problemsTabBtn">Skipped &amp; errors</button>
  </nav>
  <main>
    <section class="tab" id="tab-files">
      <div class="layout">
        <aside class="panel">
          <div class="hd">
            <h2>Folders</h2>
            <select id="treeSort" title="Sort folders">
              <option value="name">Sort: name</option>
              <option value="size">Sort: size</option>
            </select>
            <button class="btn" id="expandAll" type="button">Expand all</button>
            <button class="btn" id="collapseAll" type="button">Collapse all</button>
          </div>
          <div class="bd">
            <div class="toolbar"><button class="btn" id="allFilesBtn" type="button">Show all files</button><span class="note" id="treeNote"></span></div>
            <div id="tree" role="tree"></div>
          </div>
        </aside>
        <section class="panel">
          <div class="bd filters">
            <div class="frow">
              <label class="lbl" for="search">Search</label>
              <input type="search" id="search" placeholder="File name or path (case-insensitive)" autocomplete="off">
            </div>
            <div class="frow">
              <label class="lbl">Type</label>
              <span id="catChips" class="frow"></span>
              <details class="extpick" id="extPick">
                <summary class="btn" id="extSummary">Extensions</summary>
                <div class="extbox">
                  <div class="frow"><input type="search" id="extSearch" placeholder="Find extension" autocomplete="off" style="flex:1"><button class="btn" id="extClear" type="button">Clear</button></div>
                  <div class="extlist" id="extList"></div>
                  <div class="note" id="extNote"></div>
                </div>
              </details>
            </div>
            <div class="frow">
              <label class="lbl" for="datePreset">Modified</label>
              <select id="datePreset">
                <option value="">Any time</option>
                <option value="1">Last 24 hours</option>
                <option value="7">Last 7 days</option>
                <option value="30">Last 30 days</option>
                <option value="90">Last 90 days</option>
                <option value="365">Last year (365 days)</option>
                <option value="custom">Custom range</option>
              </select>
              <label class="note" for="dateFrom">from</label><input type="date" id="dateFrom">
              <label class="note" for="dateTo">to</label><input type="date" id="dateTo">
            </div>
            <div class="frow">
              <label class="lbl" for="minVal">Minimum size</label>
              <input type="number" id="minVal" min="0" step="any" placeholder="0">
              <select id="minUnit">
                <option value="1">bytes</option>
                <option value="1024">KB</option>
                <option value="1048576" selected>MB</option>
                <option value="1073741824">GB</option>
                <option value="1099511627776">TB</option>
              </select>
              <span class="note">(1 KB = 1024 bytes)</span>
            </div>
            <div class="frow">
              <button class="btn" id="clearAll" type="button">Clear all filters</button>
              <button class="btn primary" id="exportCsv" type="button">Export CSV</button>
              <label class="note"><input type="checkbox" id="showFull"> Show full path column</label>
              <label class="note"><input type="checkbox" id="showCreated"> Show created date</label>
            </div>
            <div id="activeFilters" aria-live="polite"></div>
          </div>
          <div class="tablebar">
            <span class="count" id="countText" aria-live="polite"></span>
            <label class="note">Rows per page
              <select id="pageSize"><option>25</option><option>50</option><option selected>100</option><option>250</option><option>500</option><option>1000</option></select>
            </label>
            <span class="pager" id="pagerTop"></span>
          </div>
          <div class="tablewrap">
            <table id="fileTable">
              <thead><tr id="headRow"></tr></thead>
              <tbody id="fileBody"></tbody>
            </table>
            <div class="empty" id="emptyMsg" hidden>No files match the current filters. <button class="link" id="emptyClear" type="button">Clear all filters</button></div>
          </div>
          <div class="tablebar"><span class="count" id="countText2"></span><span class="pager" id="pagerBottom"></span></div>
        </section>
      </div>
    </section>

    <section class="tab" id="tab-analysis" hidden>
      <p class="note">Based on the complete inventory (filters on the Files tab do not apply here). Click a folder or file to open it in the Files tab. Nothing is deleted or modified by this report.</p>
      <div class="grid2">
        <div class="panel"><div class="hd"><h2>Largest folders (total, including subfolders)</h2></div><div class="bd tablewrap" id="anFoldersTotal"></div></div>
        <div class="panel"><div class="hd"><h2>Folders with the most space in files directly inside</h2></div><div class="bd tablewrap" id="anFoldersDirect"></div></div>
        <div class="panel"><div class="hd"><h2>Largest files</h2></div><div class="bd tablewrap" id="anFiles"></div></div>
        <div class="panel"><div class="hd"><h2>File categories</h2></div><div class="bd tablewrap" id="anCats"></div></div>
        <div class="panel"><div class="hd"><h2>Extensions by total size</h2></div><div class="bd tablewrap" id="anExtSize"></div></div>
        <div class="panel"><div class="hd"><h2>Extensions by file count</h2></div><div class="bd tablewrap" id="anExtCount"></div></div>
      </div>
    </section>

    <section class="tab" id="tab-problems" hidden>
      <div class="panel">
        <div class="hd">
          <h2>Skipped folders and errors</h2>
          <select id="probKind"></select>
          <button class="btn" id="exportProblems" type="button">Export CSV</button>
        </div>
        <div class="bd">
          <p class="note" id="probIntro"></p>
          <div class="tablewrap" id="probTable"></div>
        </div>
      </div>
      <div class="panel" style="margin-top:14px">
        <div class="hd"><h2>Not included by design</h2></div>
        <div class="bd" id="notIncluded"></div>
      </div>
    </section>
  </main>
  <footer id="footer"></footer>
</div>
<script type="application/json" id="inventory-data">__INVENTORY_DATA__</script>
<script>
(function () {
  'use strict';

  function $(id) { return document.getElementById(id); }
  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text !== undefined && text !== null) e.textContent = String(text);
    return e;
  }
  function fatal(msg) { var f = $('fatal'); f.textContent = msg; f.hidden = false; }

  var D;
  try {
    D = JSON.parse($('inventory-data').textContent);
  } catch (err) {
    fatal('The inventory data in this report could not be read: ' + err.message);
    return;
  }

  // ---------------------------------------------------------------- data setup
  var META = D.meta, SEP = META.sep || '\\', CATS = D.cats;
  var NE = D.exts.length, EXT_NAME = new Array(NE), EXT_CAT = new Int32Array(NE);
  for (var x = 0; x < NE; x++) { EXT_NAME[x] = D.exts[x][0]; EXT_CAT[x] = D.exts[x][1]; }
  var NO_EXT_CAT = CATS.indexOf('No extension');

  var DR = D.dirs, ND = DR.length;
  var dParent = new Int32Array(ND), dName = new Array(ND), dMtime = new Float64Array(ND), dStatus = new Int8Array(ND),
      dDirSize = new Float64Array(ND), dDirFiles = new Float64Array(ND), dTotSize = new Float64Array(ND), dTotFiles = new Float64Array(ND),
      dSubdirs = new Int32Array(ND), dDescDirs = new Int32Array(ND), dPartial = new Uint8Array(ND), dErrType = new Array(ND),
      dChildren = new Array(ND), dRel = new Array(ND), dPre = new Int32Array(ND), dLast = new Int32Array(ND), dDepth = new Int32Array(ND);
  for (var i = 0; i < ND; i++) {
    var r = DR[i];
    dParent[i] = r[0]; dName[i] = r[1]; dMtime[i] = r[2]; dStatus[i] = r[3]; dDirSize[i] = r[4]; dDirFiles[i] = r[5];
    dTotSize[i] = r[6]; dTotFiles[i] = r[7]; dSubdirs[i] = r[8]; dDescDirs[i] = r[9]; dPartial[i] = r[10]; dErrType[i] = r[12];
    dChildren[i] = [];
  }
  for (i = 1; i < ND; i++) { if (dParent[i] >= 0 && dParent[i] < ND) dChildren[dParent[i]].push(i); }
  // Pre-order numbering: a folder's descendants occupy the range dPre[d]..dLast[d].
  var visitOrder = [];
  (function () {
    var stack = [0]; dRel[0] = '';
    while (stack.length) {
      var d = stack.pop();
      dPre[d] = visitOrder.length; visitOrder.push(d);
      var ch = dChildren[d];
      for (var k = ch.length - 1; k >= 0; k--) {
        var c = ch[k];
        dRel[c] = dRel[d] ? dRel[d] + SEP + dName[c] : dName[c];
        dDepth[c] = dDepth[d] + 1;
        stack.push(c);
      }
    }
    for (var o = visitOrder.length - 1; o >= 0; o--) {
      var dd = visitOrder[o], m = dPre[dd], cc = dChildren[dd];
      for (var q = 0; q < cc.length; q++) if (dLast[cc[q]] > m) m = dLast[cc[q]];
      dLast[dd] = m;
    }
  })();

  var FL = D.files, STRIDE = 7, N = Math.floor(FL.length / STRIDE);
  var fDir = new Int32Array(N), fName = new Array(N), fSize = new Float64Array(N), fM = new Float64Array(N),
      fC = new Float64Array(N), fAttr = new Int32Array(N), fExt = new Int32Array(N), fPre = new Int32Array(N);
  var totalBytes = 0;
  for (i = 0, x = 0; i < N; i++, x += STRIDE) {
    fDir[i] = FL[x]; fName[i] = FL[x + 1]; fSize[i] = FL[x + 2]; fM[i] = FL[x + 3]; fC[i] = FL[x + 4];
    fAttr[i] = FL[x + 5]; fExt[i] = FL[x + 6]; fPre[i] = dPre[FL[x]]; totalBytes += FL[x + 2];
  }
  D.files = null; FL = null;

  var catCount = new Float64Array(CATS.length), catSize = new Float64Array(CATS.length);
  var extCount = new Float64Array(NE), extSize = new Float64Array(NE);
  for (i = 0; i < N; i++) { var e0 = fExt[i]; extCount[e0]++; extSize[e0] += fSize[i]; catCount[EXT_CAT[e0]]++; catSize[EXT_CAT[e0]] += fSize[i]; }

  // Lazily built lower-case keys for search and sorting.
  var fLower = null, dRelLower = null, dRank = null, extRank = null;
  function ensureLower() {
    if (fLower) return;
    fLower = new Array(N);
    for (var j = 0; j < N; j++) fLower[j] = fName[j].toLowerCase();
    dRelLower = new Array(ND);
    for (j = 0; j < ND; j++) dRelLower[j] = (dRel[j] || '').toLowerCase();
  }
  function strcmp(a, b) { return a < b ? -1 : (a > b ? 1 : 0); }
  function ensureRanks() {
    ensureLower();
    if (dRank) return;
    var idx = []; for (var j = 0; j < ND; j++) idx.push(j);
    idx.sort(function (a, b) { return strcmp(dRelLower[a], dRelLower[b]) || a - b; });
    dRank = new Int32Array(ND); for (j = 0; j < ND; j++) dRank[idx[j]] = j;
    var ei = []; for (j = 0; j < NE; j++) ei.push(j);
    ei.sort(function (a, b) { return strcmp(EXT_NAME[a], EXT_NAME[b]); });
    extRank = new Int32Array(NE); for (j = 0; j < NE; j++) extRank[ei[j]] = j;
  }

  // ---------------------------------------------------------------- formatting
  var nf = new Intl.NumberFormat();
  function fmtInt(n) { return nf.format(n); }
  function fmtBytes(b) {
    if (b < 1024) return fmtInt(b) + ' B';
    var u = ['KB', 'MB', 'GB', 'TB', 'PB'], v = b / 1024, k = 0;
    while (v >= 1024 && k < u.length - 1) { v /= 1024; k++; }
    return (v >= 100 ? v.toFixed(0) : (v >= 10 ? v.toFixed(1) : v.toFixed(2))) + ' ' + u[k];
  }
  function plural(n, w) { return fmtInt(n) + ' ' + w + (n === 1 ? '' : 's'); }
  function pad(n) { return n < 10 ? '0' + n : String(n); }
  function fmtDate(sec, withSec) {
    var d = new Date(sec * 1000);
    if (isNaN(d.getTime())) return '';
    var s = d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) + ' ' + pad(d.getHours()) + ':' + pad(d.getMinutes());
    return withSec ? s + ':' + pad(d.getSeconds()) : s;
  }
  function attrStr(a) {
    var s = '';
    if (a & 1) s += 'R'; if (a & 2) s += 'H'; if (a & 4) s += 'S'; if (a & 32) s += 'A';
    if (a & 1024) s += 'L'; if (a & 2048) s += 'C'; if (a & 16384) s += 'E'; if ((a & 4096) || (a & 0x40000) || (a & 0x400000)) s += 'O';
    return s;
  }
  var ATTR_HELP = 'R=read-only H=hidden S=system A=archive L=reparse point/link C=compressed E=encrypted O=offline/cloud placeholder';
  function relPath(i) { var r = dRel[fDir[i]]; return r ? r + SEP + fName[i] : fName[i]; }
  function joinRoot(rel) { var root = META.root; if (!rel) return root; return (root.charAt(root.length - 1) === SEP ? root : root + SEP) + rel; }
  function fullPath(i) { return joinRoot(relPath(i)); }
  function folderLabel(d) { return d === 0 ? META.root : dRel[d]; }
  function extLabel(e) { return EXT_NAME[e] ? '.' + EXT_NAME[e] : '(no extension)'; }
  var STATUS_TEXT = { 1: 'Partially listed', 2: 'Skipped', 3: 'Not scanned' };
  function dirStatusText(d) {
    if (dStatus[d] === 2 && dErrType[d] === 'AccessDenied') return 'Access denied';
    return STATUS_TEXT[dStatus[d]] || '';
  }

  // ---------------------------------------------------------------- header and cards
  var C = META.counts, complete = META.status === 'Complete';
  document.title = 'Folder Inventory - ' + META.root;
  $('mRoot').textContent = META.root;
  $('mGen').textContent = META.generated;
  $('mDur').textContent = META.durationSec + ' s';
  var o = META.options || {};
  $('mOpt').textContent = 'workers ' + o.throttleLimit + ', hidden ' + (o.includeHidden ? 'included' : 'excluded') + ', system ' +
    (o.includeSystem ? 'included' : 'excluded') + ', links ' + (o.followReparsePoints ? 'followed' : 'not followed') +
    (o.timeoutMinutes ? ', time limit ' + o.timeoutMinutes + ' min' : '');
  var sb = $('statusBadge');
  sb.textContent = complete ? 'Scan complete' : 'Partial scan';
  sb.className = 'badge ' + (complete ? 'ok' : 'warn');
  if (!complete) {
    var bn = $('banner');
    bn.textContent = 'Partial scan: ' + fmtInt(C.skippedFolders) + ' folder(s) could not be read, ' + fmtInt(C.partialFolders) +
      ' folder(s) were only partly listed and ' + fmtInt(C.problems) + ' problem(s) were recorded' + (META.timedOut ? ' (the time limit was reached)' : '') +
      '. Sizes and counts are lower bounds. See the "Skipped & errors" tab.';
    bn.hidden = false;
  }
  $('problemsTabBtn').textContent = 'Skipped & errors (' + fmtInt(C.problems) + ')';

  function card(k, v, s, title) {
    var c = el('div', 'card'); c.appendChild(el('div', 'k', k));
    var vv = el('div', 'v', v); if (title) vv.title = title; c.appendChild(vv);
    if (s) { var ss = el('div', 's', s); ss.title = s; c.appendChild(ss); }
    $('cards').appendChild(c);
    return c;
  }
  var largest = -1;
  for (i = 0; i < N; i++) if (largest < 0 || fSize[i] > fSize[largest]) largest = i;
  var topExt = -1;
  for (x = 0; x < NE; x++) if (topExt < 0 || extCount[x] > extCount[topExt]) topExt = x;
  card('Total size', fmtBytes(totalBytes), fmtInt(totalBytes) + ' bytes' + (complete ? '' : ' (lower bound)'));
  card('Files', fmtInt(N), C.excludedFiles ? fmtInt(C.excludedFiles) + ' hidden/system excluded' : 'all accessible files');
  card('Folders', fmtInt(dDescDirs[0]), 'subfolders under the root');
  if (largest >= 0) card('Largest file', fmtBytes(fSize[largest]), relPath(largest), fName[largest]);
  else card('Largest file', '-', 'no files');
  if (topExt >= 0) card('Most common extension', extLabel(topExt), fmtInt(extCount[topExt]) + ' files, ' + fmtBytes(extSize[topExt]));
  else card('Most common extension', '-', 'no files');
  var skc = card('Skipped folders', fmtInt(C.skippedFolders), fmtInt(C.problems) + ' problem(s) recorded in total');
  if (C.skippedFolders) skc.querySelector('.v').style.color = 'var(--warn)';
  var stc = card('Scan status', complete ? 'Complete' : 'Partial', complete ? 'every accessible item was read' : 'some items could not be read');
  stc.querySelector('.v').style.color = complete ? 'var(--ok)' : 'var(--warn)';

  // ---------------------------------------------------------------- tabs
  var analysisBuilt = false, problemsBuilt = false;
  function showTab(name) {
    var btns = document.querySelectorAll('nav.tabs button');
    for (var b = 0; b < btns.length; b++) btns[b].setAttribute('aria-selected', btns[b].getAttribute('data-tab') === name ? 'true' : 'false');
    $('tab-files').hidden = name !== 'files';
    $('tab-analysis').hidden = name !== 'analysis';
    $('tab-problems').hidden = name !== 'problems';
    if (name === 'analysis' && !analysisBuilt) { analysisBuilt = true; buildAnalysis(); }
    if (name === 'problems' && !problemsBuilt) { problemsBuilt = true; buildProblems(); }
  }
  Array.prototype.forEach.call(document.querySelectorAll('nav.tabs button'), function (b) {
    b.addEventListener('click', function () { showTab(b.getAttribute('data-tab')); });
  });

  // ---------------------------------------------------------------- filter state
  var state = {
    q: '', folder: -1, cats: new Set(), exts: new Set(), preset: '', from: '', to: '', minVal: '', minUnit: 1048576,
    sortKey: 'size', sortDir: -1, page: 0, pageSize: 100, showFull: false, showCreated: false
  };
  var view = new Int32Array(0), viewBytes = 0;

  function parseDay(s, endOfDay) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || '');
    if (!m) return null;
    var d = endOfDay ? new Date(+m[1], +m[2] - 1, +m[3], 23, 59, 59, 999) : new Date(+m[1], +m[2] - 1, +m[3], 0, 0, 0, 0);
    return d.getTime() / 1000;
  }
  function dateRange() {
    if (!state.preset) return null;
    if (state.preset === 'custom') {
      var lo = parseDay(state.from, false), hi = parseDay(state.to, true);
      if (lo === null && hi === null) return null;
      return [lo === null ? -Infinity : lo, hi === null ? Infinity : hi];
    }
    return [Date.now() / 1000 - Number(state.preset) * 86400, Infinity];
  }
  function minBytes() {
    var v = parseFloat(state.minVal);
    if (!isFinite(v) || v <= 0) return 0;
    return v * state.minUnit;
  }

  function computeView() {
    var q = state.q.trim().toLowerCase();
    if (SEP === '\\') q = q.replace(/\//g, '\\');
    var hasQ = q.length > 0, qSep = q.indexOf(SEP) >= 0, dMatch = null;
    if (hasQ) {
      ensureLower();
      dMatch = new Uint8Array(ND);
      for (var d = 0; d < ND; d++) if (dRelLower[d] && dRelLower[d].indexOf(q) >= 0) dMatch[d] = 1;
    }
    var folder = state.folder, lo = folder >= 0 ? dPre[folder] : 0, hi = folder >= 0 ? dLast[folder] : 0;
    var dr = dateRange(), dlo = dr ? dr[0] : 0, dhi = dr ? dr[1] : 0;
    var mb = minBytes();
    var hasType = state.cats.size > 0 || state.exts.size > 0, typeOk = null;
    if (hasType) {
      typeOk = new Uint8Array(NE);
      for (var e = 0; e < NE; e++) if (state.cats.has(EXT_CAT[e]) || state.exts.has(e)) typeOk[e] = 1;
    }
    var out = new Int32Array(N), n = 0, bytes = 0;
    for (var j = 0; j < N; j++) {
      if (mb && fSize[j] < mb) continue;
      if (folder >= 0) { var p = fPre[j]; if (p < lo || p > hi) continue; }
      if (dr) { var mt = fM[j]; if (mt < dlo || mt > dhi) continue; }
      if (hasType && !typeOk[fExt[j]]) continue;
      if (hasQ && fLower[j].indexOf(q) < 0 && !dMatch[fDir[j]]) {
        if (!qSep) continue;
        var rl = dRelLower[fDir[j]];
        if (((rl ? rl + SEP : '') + fLower[j]).indexOf(q) < 0) continue;
      }
      out[n++] = j; bytes += fSize[j];
    }
    view = out.slice(0, n); viewBytes = bytes;
    sortView();
  }

  function comparator() {
    var dir = state.sortDir;
    switch (state.sortKey) {
      case 'name': ensureLower(); return function (a, b) { return strcmp(fLower[a], fLower[b]) * dir || a - b; };
      case 'path': ensureRanks(); return function (a, b) { return ((dRank[fDir[a]] - dRank[fDir[b]]) || strcmp(fLower[a], fLower[b])) * dir || a - b; };
      case 'ext': ensureRanks(); return function (a, b) { return ((extRank[fExt[a]] - extRank[fExt[b]]) || strcmp(fLower[a], fLower[b])) * dir || a - b; };
      case 'mtime': return function (a, b) { return (fM[a] - fM[b]) * dir || a - b; };
      case 'ctime': return function (a, b) { return (fC[a] - fC[b]) * dir || a - b; };
      case 'attr': return function (a, b) { return (fAttr[a] - fAttr[b]) * dir || a - b; };
      default: return function (a, b) { return (fSize[a] - fSize[b]) * dir || a - b; };
    }
  }
  function sortView() { view.sort(comparator()); }

  // ---------------------------------------------------------------- table
  var COLS = [
    { key: 'name', label: 'Name' },
    { key: 'path', label: 'Relative path' },
    { key: 'ext', label: 'Type' },
    { key: 'size', label: 'Size', num: true },
    { key: 'mtime', label: 'Last modified' },
    { key: 'ctime', label: 'Created', opt: 'showCreated' },
    { key: 'attr', label: 'Attr', title: ATTR_HELP },
    { key: 'full', label: 'Full path', opt: 'showFull', nosort: true }
  ];
  function visibleCols() { return COLS.filter(function (c) { return !c.opt || state[c.opt]; }); }
  function renderHead() {
    var tr = $('headRow'); tr.textContent = '';
    visibleCols().forEach(function (c) {
      var th = el('th', (c.num ? 'num ' : '') + (c.nosort ? '' : 'sortable'));
      if (c.title) th.title = c.title;
      if (!c.nosort) {
        th.setAttribute('tabindex', '0');
        th.setAttribute('aria-sort', state.sortKey === c.key ? (state.sortDir > 0 ? 'ascending' : 'descending') : 'none');
        var go = function () {
          if (state.sortKey === c.key) state.sortDir = -state.sortDir;
          else { state.sortKey = c.key; state.sortDir = (c.key === 'size' || c.key === 'mtime' || c.key === 'ctime') ? -1 : 1; }
          state.page = 0; sortView(); renderHead(); renderTable();
        };
        th.addEventListener('click', go);
        th.addEventListener('keydown', function (ev) { if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); go(); } });
      }
      th.appendChild(document.createTextNode(c.label + ' '));
      th.appendChild(el('span', 'arrow', state.sortKey === c.key ? (state.sortDir > 0 ? '\u25B2' : '\u25BC') : ''));
      tr.appendChild(th);
    });
  }

  function pager(container) {
    container.textContent = '';
    var pages = Math.max(1, Math.ceil(view.length / state.pageSize));
    function b(label, target, disabled, title) {
      var bt = el('button', 'btn', label); bt.type = 'button'; bt.disabled = disabled; if (title) bt.title = title;
      bt.addEventListener('click', function () { state.page = target; renderTable(); });
      container.appendChild(bt);
    }
    b('\u00AB', 0, state.page === 0, 'First page');
    b('\u2039', state.page - 1, state.page === 0, 'Previous page');
    var inp = el('input'); inp.type = 'number'; inp.min = 1; inp.max = pages; inp.value = state.page + 1; inp.title = 'Go to page';
    inp.addEventListener('change', function () {
      var v = parseInt(inp.value, 10);
      if (isFinite(v)) { state.page = Math.min(pages, Math.max(1, v)) - 1; renderTable(); }
    });
    container.appendChild(inp);
    container.appendChild(el('span', 'note', 'of ' + fmtInt(pages)));
    b('\u203A', state.page + 1, state.page >= pages - 1, 'Next page');
    b('\u00BB', pages - 1, state.page >= pages - 1, 'Last page');
  }

  function renderTable() {
    var total = view.length, ps = state.pageSize, pages = Math.max(1, Math.ceil(total / ps));
    if (state.page >= pages) state.page = pages - 1;
    if (state.page < 0) state.page = 0;
    var start = state.page * ps, end = Math.min(total, start + ps);
    var body = $('fileBody'); body.textContent = '';
    var cols = visibleCols(), frag = document.createDocumentFragment();
    for (var k = start; k < end; k++) {
      var i = view[k], tr = el('tr');
      for (var c = 0; c < cols.length; c++) {
        var key = cols[c].key, td;
        if (key === 'name') { td = el('td', 'name', fName[i]); }
        else if (key === 'path') {
          td = el('td', 'path');
          var fd = fDir[i];
          if (dRel[fd]) {
            var fl = el('button', 'link folder', dRel[fd] + SEP); fl.type = 'button'; fl.title = 'Show only this folder';
            fl.setAttribute('data-dir', fd);
            td.appendChild(fl);
          }
          td.appendChild(document.createTextNode(fName[i]));
        }
        else if (key === 'ext') {
          td = el('td', 'type'); td.appendChild(el('div', null, extLabel(fExt[i]))); td.appendChild(el('div', 'cat', CATS[EXT_CAT[fExt[i]]]));
        }
        else if (key === 'size') { td = el('td', 'num', fmtBytes(fSize[i])); td.title = fmtInt(fSize[i]) + ' bytes'; }
        else if (key === 'mtime') { td = el('td', 'date', fmtDate(fM[i])); td.title = fmtDate(fM[i], true); }
        else if (key === 'ctime') { td = el('td', 'date', fmtDate(fC[i])); td.title = fmtDate(fC[i], true); }
        else if (key === 'attr') { td = el('td', 'attr', attrStr(fAttr[i])); td.title = ATTR_HELP; }
        else { td = el('td', 'full', fullPath(i)); }
        tr.appendChild(td);
      }
      frag.appendChild(tr);
    }
    body.appendChild(frag);
    $('emptyMsg').hidden = total !== 0;
    $('fileTable').style.display = total === 0 ? 'none' : '';
    var txt = total === 0 ? 'No matching files (of ' + fmtInt(N) + ' total)' :
      'Showing ' + fmtInt(start + 1) + '\u2013' + fmtInt(end) + ' of ' + fmtInt(total) + ' matching files (' + fmtBytes(viewBytes) + ')' +
      (total !== N ? ' \u00B7 ' + fmtInt(N) + ' files in inventory' : '');
    $('countText').textContent = txt; $('countText2').textContent = txt;
    pager($('pagerTop')); pager($('pagerBottom'));
  }
  $('fileBody').addEventListener('click', function (ev) {
    var t = ev.target;
    if (t && t.getAttribute && t.getAttribute('data-dir') !== null) setFolder(parseInt(t.getAttribute('data-dir'), 10), true);
  });

  // ---------------------------------------------------------------- active filter chips
  function renderActive() {
    var box = $('activeFilters'); box.textContent = '';
    function chip(text, onRemove) {
      var c = el('span', 'af'); var s = el('span', null, text); s.title = text; c.appendChild(s);
      var b = el('button', null, '\u00D7'); b.type = 'button'; b.title = 'Remove filter'; b.setAttribute('aria-label', 'Remove filter: ' + text);
      b.addEventListener('click', onRemove); c.appendChild(b); box.appendChild(c);
    }
    if (state.folder >= 0) chip('Folder: ' + folderLabel(state.folder) + ' (and subfolders)', function () { setFolder(-1); });
    if (state.q.trim()) chip('Search: ' + state.q.trim(), function () { $('search').value = ''; state.q = ''; refresh(); });
    state.cats.forEach(function (c) { chip('Category: ' + CATS[c], function () { state.cats.delete(c); syncTypeUi(); refresh(); }); });
    state.exts.forEach(function (e) { chip('Extension: ' + extLabel(e), function () { state.exts.delete(e); syncTypeUi(); refresh(); }); });
    var dr = dateRange();
    if (dr) {
      var label = state.preset === 'custom'
        ? 'Modified: ' + (state.from || 'any') + ' to ' + (state.to || 'any') + ' (inclusive)'
        : 'Modified: ' + $('datePreset').options[$('datePreset').selectedIndex].text;
      chip(label, function () { state.preset = ''; state.from = ''; state.to = ''; syncDateUi(); refresh(); });
    }
    var mb = minBytes();
    if (mb) chip('Size \u2265 ' + state.minVal + ' ' + $('minUnit').options[$('minUnit').selectedIndex].text + ' (' + fmtInt(Math.ceil(mb)) + ' bytes)',
      function () { state.minVal = ''; $('minVal').value = ''; refresh(); });
  }

  function refresh() {
    state.page = 0;
    computeView();
    renderActive();
    renderTable();
  }

  // ---------------------------------------------------------------- type filters
  var catBtns = [];
  (function () {
    var box = $('catChips');
    CATS.forEach(function (name, ci) {
      if (!catCount[ci]) return;
      var b = el('button', 'chip', ci === NO_EXT_CAT ? 'Extensionless files' : name); b.type = 'button';
      b.appendChild(el('span', 'n', fmtInt(catCount[ci])));
      b.setAttribute('aria-pressed', 'false');
      b.title = ci === NO_EXT_CAT ? 'Include files that have no extension' : 'Files in category ' + name;
      b.addEventListener('click', function () {
        if (state.cats.has(ci)) state.cats.delete(ci); else state.cats.add(ci);
        syncTypeUi(); refresh();
      });
      box.appendChild(b); catBtns.push([ci, b]);
    });
  })();
  var extOrder = []; for (x = 0; x < NE; x++) extOrder.push(x);
  extOrder.sort(function (a, b) { return extCount[b] - extCount[a] || strcmp(EXT_NAME[a], EXT_NAME[b]); });
  function renderExtList() {
    var list = $('extList'); list.textContent = '';
    var q = $('extSearch').value.trim().toLowerCase().replace(/^\./, ''), shown = 0, matches = 0, LIMIT = 300;
    for (var k = 0; k < extOrder.length; k++) {
      var e = extOrder[k], label = extLabel(e);
      if (q && EXT_NAME[e].indexOf(q) < 0 && !(q === 'none' && !EXT_NAME[e])) continue;
      matches++;
      if (shown >= LIMIT && !state.exts.has(e)) continue;
      shown++;
      var lb = el('label'), cb = el('input'); cb.type = 'checkbox'; cb.checked = state.exts.has(e);
      (function (ee, box) {
        box.addEventListener('change', function () { if (box.checked) state.exts.add(ee); else state.exts.delete(ee); syncTypeUi(); refresh(); });
      })(e, cb);
      lb.appendChild(cb); lb.appendChild(el('span', null, label)); lb.appendChild(el('span', 'n', fmtInt(extCount[e]) + ' \u00B7 ' + fmtBytes(extSize[e])));
      list.appendChild(lb);
    }
    $('extNote').textContent = matches > shown ? 'Showing ' + shown + ' of ' + matches + ' extensions; type to narrow the list.' : (matches ? '' : 'No extensions match.');
  }
  function syncTypeUi() {
    catBtns.forEach(function (p) { p[1].setAttribute('aria-pressed', state.cats.has(p[0]) ? 'true' : 'false'); });
    $('extSummary').textContent = state.exts.size ? 'Extensions (' + state.exts.size + ' selected)' : 'Extensions';
    if ($('extPick').open) renderExtList();
  }
  $('extPick').addEventListener('toggle', function () { if ($('extPick').open) { renderExtList(); $('extSearch').focus(); } });
  $('extSearch').addEventListener('input', renderExtList);
  $('extClear').addEventListener('click', function () { state.exts.clear(); syncTypeUi(); refresh(); });
  document.addEventListener('click', function (ev) { var p = $('extPick'); if (p.open && !p.contains(ev.target)) p.open = false; });

  // ---------------------------------------------------------------- other filter controls
  var searchTimer = null;
  $('search').addEventListener('input', function () {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(function () { state.q = $('search').value; refresh(); }, 250);
  });
  function syncDateUi() {
    $('datePreset').value = state.preset;
    $('dateFrom').value = state.from; $('dateTo').value = state.to;
    var custom = state.preset === 'custom';
    $('dateFrom').disabled = !custom && !!state.preset; $('dateTo').disabled = !custom && !!state.preset;
  }
  $('datePreset').addEventListener('change', function () {
    state.preset = $('datePreset').value;
    if (state.preset !== 'custom') { state.from = ''; state.to = ''; }
    syncDateUi(); refresh();
  });
  function onDateInput() {
    state.from = $('dateFrom').value; state.to = $('dateTo').value;
    state.preset = (state.from || state.to) ? 'custom' : '';
    syncDateUi(); refresh();
  }
  $('dateFrom').addEventListener('change', onDateInput);
  $('dateTo').addEventListener('change', onDateInput);
  $('minVal').addEventListener('input', function () { state.minVal = $('minVal').value; refresh(); });
  $('minUnit').addEventListener('change', function () { state.minUnit = Number($('minUnit').value); refresh(); });
  $('pageSize').addEventListener('change', function () { state.pageSize = parseInt($('pageSize').value, 10) || 100; state.page = 0; renderTable(); });
  $('showFull').addEventListener('change', function () { state.showFull = $('showFull').checked; renderHead(); renderTable(); });
  $('showCreated').addEventListener('change', function () { state.showCreated = $('showCreated').checked; renderHead(); renderTable(); });
  function clearAll() {
    state.q = ''; state.folder = -1; state.cats.clear(); state.exts.clear(); state.preset = ''; state.from = ''; state.to = ''; state.minVal = '';
    $('search').value = ''; $('minVal').value = '';
    syncDateUi(); syncTypeUi(); renderTree(); refresh();
  }
  $('clearAll').addEventListener('click', clearAll);
  $('emptyClear').addEventListener('click', clearAll);

  // ---------------------------------------------------------------- folder tree (lazy: only expanded folders are rendered)
  var expanded = new Set([0]), treeSort = 'name', childLimit = {}, TREE_CAP = 4000, CHILD_PAGE = 500;
  function sortedChildren(d) {
    var ch = dChildren[d].slice();
    if (treeSort === 'size') ch.sort(function (a, b) { return dTotSize[b] - dTotSize[a] || strcmp(dName[a].toLowerCase(), dName[b].toLowerCase()); });
    else ch.sort(function (a, b) { return strcmp(dName[a].toLowerCase(), dName[b].toLowerCase()) || a - b; });
    return ch;
  }
  function treeRow(d) {
    var row = el('div', 'trow' + (state.folder === d || (state.folder < 0 && d === 0) ? ' sel' : ''));
    row.setAttribute('role', 'treeitem');
    row.style.paddingLeft = (dDepth[d] * 14 + 4) + 'px';
    if (dChildren[d].length) {
      var isOpen = expanded.has(d);
      row.setAttribute('aria-expanded', isOpen ? 'true' : 'false');
      var t = el('button', 'ttog', isOpen ? '\u25BE' : '\u25B8'); t.type = 'button'; t.title = isOpen ? 'Collapse' : 'Expand';
      t.setAttribute('data-toggle', d);
      row.appendChild(t);
    } else row.appendChild(el('span', 'tspace'));
    var nm = el('button', 'tname', dName[d]); nm.type = 'button'; nm.setAttribute('data-select', d);
    nm.title = (d === 0 ? META.root : dRel[d]) + '\nDirect: ' + fmtInt(dDirFiles[d]) + ' files, ' + fmtBytes(dDirSize[d]) +
      '\nTotal: ' + fmtInt(dTotFiles[d]) + ' files, ' + fmtBytes(dTotSize[d]) + ' (' + fmtInt(dTotSize[d]) + ' bytes)' +
      '\nSubfolders: ' + fmtInt(dSubdirs[d]) + ' direct, ' + fmtInt(dDescDirs[d]) + ' total' +
      (dMtime[d] ? '\nModified: ' + fmtDate(dMtime[d], true) : '') + (dPartial[d] ? '\nTotals are partial (some content could not be read)' : '');
    row.appendChild(nm);
    var st = dirStatusText(d);
    if (st) row.appendChild(el('span', 'badge ' + (dStatus[d] === 1 ? 'warn' : 'bad'), st));
    else if (dPartial[d]) row.appendChild(el('span', 'badge warn', 'partial'));
    if (!st && dTotFiles[d] === 0) row.appendChild(el('span', 'badge mute', dDescDirs[d] === 0 ? 'empty' : 'no files'));
    row.appendChild(el('span', 'tstat', (dPartial[d] ? '\u2265 ' : '') + fmtBytes(dTotSize[d]) + ' \u00B7 ' + plural(dTotFiles[d], 'file') + ' \u00B7 ' + plural(dDescDirs[d], 'folder')));
    return row;
  }
  function renderTree() {
    var tree = $('tree'), frag = document.createDocumentFragment(), count = 0;
    tree.textContent = '';
    (function walk(d) {
      frag.appendChild(treeRow(d)); count++;
      if (!expanded.has(d) || !dChildren[d].length) return;
      var ch = sortedChildren(d), lim = childLimit[d] || CHILD_PAGE;
      for (var k = 0; k < ch.length && k < lim; k++) walk(ch[k]);
      if (ch.length > lim) {
        var more = el('div', 'trow tmore'); more.style.paddingLeft = ((dDepth[d] + 1) * 14 + 22) + 'px';
        var mb = el('button', 'link', 'Show ' + fmtInt(Math.min(CHILD_PAGE, ch.length - lim)) + ' more of ' + fmtInt(ch.length - lim) + ' remaining folders');
        mb.type = 'button'; mb.setAttribute('data-more', d); more.appendChild(mb); frag.appendChild(more);
      }
    })(0);
    tree.appendChild(frag);
  }
  $('tree').addEventListener('click', function (ev) {
    var t = ev.target; if (!t || !t.getAttribute) return;
    var v;
    if ((v = t.getAttribute('data-toggle')) !== null) { v = +v; if (expanded.has(v)) expanded.delete(v); else expanded.add(v); $('treeNote').textContent = ''; renderTree(); }
    else if ((v = t.getAttribute('data-select')) !== null) setFolder(+v, false);
    else if ((v = t.getAttribute('data-more')) !== null) { v = +v; childLimit[v] = (childLimit[v] || CHILD_PAGE) + CHILD_PAGE; renderTree(); }
  });
  $('treeSort').addEventListener('change', function () { treeSort = $('treeSort').value; renderTree(); });
  $('expandAll').addEventListener('click', function () {
    // Breadth-first expansion, capped so the page stays responsive on huge trees.
    var queue = [0], rendered = 1, capped = false;
    expanded = new Set();
    while (queue.length) {
      var d = queue.shift();
      if (!dChildren[d].length) continue;
      if (rendered + Math.min(dChildren[d].length, CHILD_PAGE) > TREE_CAP && d !== 0) { capped = true; continue; }
      expanded.add(d); rendered += Math.min(dChildren[d].length, CHILD_PAGE);
      var ch = sortedChildren(d).slice(0, childLimit[d] || CHILD_PAGE);
      for (var k = 0; k < ch.length; k++) queue.push(ch[k]);
    }
    $('treeNote').textContent = capped ? 'Large tree: expanded the first ' + fmtInt(rendered) + ' folders. Expand deeper folders individually.' : '';
    renderTree();
  });
  $('collapseAll').addEventListener('click', function () { expanded = new Set([0]); $('treeNote').textContent = ''; renderTree(); });
  $('allFilesBtn').addEventListener('click', function () { setFolder(-1); });

  function setFolder(d, reveal) {
    state.folder = (d > 0 && d < ND) ? d : -1;
    if (state.folder > 0) { var p = dParent[state.folder]; while (p >= 0) { expanded.add(p); p = dParent[p]; } }
    renderTree();
    refresh();
    if (reveal) {
      var row = document.querySelector('#tree .trow.sel');
      if (row && row.scrollIntoView) row.scrollIntoView({ block: 'nearest' });
    }
  }

  // ---------------------------------------------------------------- CSV export
  function csvText(v) {
    var s = v === null || v === undefined ? '' : String(v);
    // Formula-injection guard for spreadsheet programs.
    if (/^[=+\-@\t\r]/.test(s)) s = "'" + s;
    return '"' + s.replace(/"/g, '""') + '"';
  }
  function csvNum(v) { return '"' + String(v) + '"'; }
  function stamp() { var d = new Date(); return d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) + '-' + pad(d.getHours()) + pad(d.getMinutes()); }
  function download(parts, filename) {
    var blob = new Blob(parts, { type: 'text/csv;charset=utf-8' });
    if (window.navigator && window.navigator.msSaveOrOpenBlob) { window.navigator.msSaveOrOpenBlob(blob, filename); return; }
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = filename; a.style.display = 'none';
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
  }
  function buildFilesCsv(indexes) {
    var parts = ['\uFEFF' + ['Name', 'RelativePath', 'FullPath', 'Folder', 'Extension', 'Category', 'SizeBytes', 'Modified', 'Created', 'Attributes'].map(csvText).join(',') + '\r\n'];
    var chunk = [];
    for (var k = 0; k < indexes.length; k++) {
      var i = indexes[k];
      chunk.push([csvText(fName[i]), csvText(relPath(i)), csvText(fullPath(i)), csvText(joinRoot(dRel[fDir[i]])), csvText(EXT_NAME[fExt[i]]),
        csvText(CATS[EXT_CAT[fExt[i]]]), csvNum(fSize[i]), csvText(fmtDate(fM[i], true)), csvText(fmtDate(fC[i], true)), csvText(attrStr(fAttr[i]))].join(','));
      if (chunk.length === 5000) { parts.push(chunk.join('\r\n') + '\r\n'); chunk = []; }
    }
    if (chunk.length) parts.push(chunk.join('\r\n') + '\r\n');
    return parts;
  }
  $('exportCsv').addEventListener('click', function () {
    download(buildFilesCsv(view), 'FolderInventory-' + stamp() + '-' + view.length + '-files.csv');
  });

  // ---------------------------------------------------------------- analysis
  function table(container, headers, rows, numCols) {
    container.textContent = '';
    if (!rows.length) { container.appendChild(el('div', 'empty', 'Nothing to show.')); return; }
    var t = el('table'), th = el('thead'), tr = el('tr');
    headers.forEach(function (h, k) { tr.appendChild(el('th', numCols.indexOf(k) >= 0 ? 'num' : '', h)); });
    th.appendChild(tr); t.appendChild(th);
    var tb = el('tbody');
    rows.forEach(function (cells) {
      var r = el('tr');
      cells.forEach(function (c, k) {
        var td = el('td', numCols.indexOf(k) >= 0 ? 'num' : '');
        if (c instanceof Node) td.appendChild(c); else td.textContent = c;
        r.appendChild(td);
      });
      tb.appendChild(r);
    });
    t.appendChild(tb); container.appendChild(t);
  }
  function bar(frac) { var b = el('div', 'bar'), i2 = el('i'); i2.style.width = Math.max(0, Math.min(100, frac * 100)).toFixed(1) + '%'; b.appendChild(i2); return b; }
  function folderLink(d) {
    var b = el('button', 'link', d === 0 ? META.root : dRel[d]); b.type = 'button';
    b.addEventListener('click', function () { showTab('files'); setFolder(d, true); });
    return b;
  }
  function topK(n, k, score) {
    var idx = []; for (var j = 0; j < n; j++) if (score(j) > 0) idx.push(j);
    idx.sort(function (a, b) { return score(b) - score(a) || a - b; });
    return idx.slice(0, k);
  }
  function buildAnalysis() {
    var TOP = 25, denom = totalBytes || 1;
    table($('anFoldersTotal'), ['Folder', 'Total size', 'Files', 'Share'],
      topK(ND, TOP + 1, function (d) { return d === 0 ? 0 : dTotSize[d]; }).slice(0, TOP).map(function (d) {
        var s = el('span', null, (dPartial[d] ? '\u2265 ' : '') + fmtBytes(dTotSize[d])); s.title = fmtInt(dTotSize[d]) + ' bytes';
        return [folderLink(d), s, fmtInt(dTotFiles[d]), bar(dTotSize[d] / denom)];
      }), [1, 2]);
    table($('anFoldersDirect'), ['Folder', 'Size of direct files', 'Direct files', 'Share'],
      topK(ND, TOP, function (d) { return dDirSize[d]; }).map(function (d) {
        var s = el('span', null, fmtBytes(dDirSize[d])); s.title = fmtInt(dDirSize[d]) + ' bytes';
        return [folderLink(d), s, fmtInt(dDirFiles[d]), bar(dDirSize[d] / denom)];
      }), [1, 2]);
    table($('anFiles'), ['File', 'Folder', 'Size', 'Modified'],
      topK(N, TOP, function (j) { return fSize[j]; }).map(function (j) {
        var nm = el('span', null, fName[j]); nm.title = fullPath(j);
        var s = el('span', null, fmtBytes(fSize[j])); s.title = fmtInt(fSize[j]) + ' bytes';
        return [nm, folderLink(fDir[j]), s, fmtDate(fM[j])];
      }), [2]);
    var maxCount = 1; for (var c = 0; c < CATS.length; c++) if (catCount[c] > maxCount) maxCount = catCount[c];
    table($('anCats'), ['Category', 'Files', 'By count', 'Total size', 'By size'],
      CATS.map(function (n2, ci) { return ci; }).filter(function (ci) { return catCount[ci] > 0; })
        .sort(function (a, b) { return catSize[b] - catSize[a]; })
        .map(function (ci) {
          var b = el('button', 'link', CATS[ci]); b.type = 'button';
          b.addEventListener('click', function () { state.cats = new Set([ci]); state.exts.clear(); syncTypeUi(); showTab('files'); refresh(); });
          return [b, fmtInt(catCount[ci]), bar(catCount[ci] / (N || 1)), fmtBytes(catSize[ci]), bar(catSize[ci] / denom)];
        }), [1, 3]);
    function extRows(scoreArr, other) {
      return topK(NE, TOP, function (e) { return scoreArr[e]; }).map(function (e) {
        var b = el('button', 'link', extLabel(e)); b.type = 'button';
        b.addEventListener('click', function () { state.exts = new Set([e]); state.cats.clear(); syncTypeUi(); showTab('files'); refresh(); });
        return [b, CATS[EXT_CAT[e]], other === 'size' ? fmtBytes(extSize[e]) : fmtInt(extCount[e]),
          other === 'size' ? fmtInt(extCount[e]) : fmtBytes(extSize[e]), bar(other === 'size' ? extSize[e] / denom : extCount[e] / (N || 1))];
      });
    }
    table($('anExtSize'), ['Extension', 'Category', 'Total size', 'Files', 'Share'], extRows(extSize, 'size'), [2, 3]);
    table($('anExtCount'), ['Extension', 'Category', 'Files', 'Total size', 'Share'], extRows(extCount, 'count'), [2, 3]);
  }

  // ---------------------------------------------------------------- problems
  var P = D.problems || [];
  function buildProblems() {
    var kinds = {}; P.forEach(function (p) { kinds[p[3]] = (kinds[p[3]] || 0) + 1; });
    var sel = $('probKind'); sel.textContent = '';
    var all = el('option', null, 'All (' + fmtInt(P.length) + ')'); all.value = ''; sel.appendChild(all);
    Object.keys(kinds).sort().forEach(function (k) { var op = el('option', null, k + ' (' + fmtInt(kinds[k]) + ')'); op.value = k; sel.appendChild(op); });
    $('probIntro').textContent = P.length
      ? fmtInt(C.skippedFolders) + ' folder(s) could not be scanned and ' + fmtInt(C.partialFolders) + ' were only partly listed. Their contents are missing from the inventory, and the totals of every folder above them are lower bounds (marked \u2265). The tool never changes permissions or ownership to get around access restrictions.'
      : 'No folders were skipped and no errors occurred.';
    function render() {
      var k = sel.value, rows = [], LIMIT = 2000, matched = 0;
      for (var j = 0; j < P.length; j++) {
        if (k && P[j][3] !== k) continue;
        matched++;
        if (rows.length < LIMIT) { var pth = el('span', 'mono', P[j][0] || '(no path)'); rows.push([P[j][3], P[j][1], pth, P[j][2]]); }
      }
      table($('probTable'), ['Kind', 'Type', 'Path', 'Details'], rows, []);
      if (matched > LIMIT) $('probTable').appendChild(el('p', 'note', 'Showing the first ' + LIMIT + ' of ' + fmtInt(matched) + '. Export CSV for the full list.'));
    }
    sel.addEventListener('change', render);
    render();
    $('exportProblems').addEventListener('click', function () {
      var parts = ['\uFEFF"Kind","Type","Path","Details"\r\n'];
      P.forEach(function (p) { parts.push([csvText(p[3]), csvText(p[1]), csvText(p[0]), csvText(p[2])].join(',') + '\r\n'); });
      download(parts, 'FolderInventory-problems-' + stamp() + '.csv');
    });
    var ni = $('notIncluded'); ni.textContent = '';
    ni.appendChild(el('p', null, fmtInt(C.excludedFiles) + ' hidden/system file(s) and ' + fmtInt(C.excludedFolders) + ' hidden/system folder(s) were excluded by the scan options' +
      ((C.excludedFiles || C.excludedFolders) ? ' (re-run with -IncludeHidden and/or -IncludeSystem to include them).' : '.')));
    var R = D.reparse || [];
    ni.appendChild(el('p', null, fmtInt(R.length) + ' junction(s)/symbolic link(s) to folders were not followed' + (R.length ? ' (re-run with -FollowReparsePoints to follow them):' : '.')));
    if (R.length) {
      var ul = el('ul', 'mono');
      R.slice(0, 1000).forEach(function (pth) { ul.appendChild(el('li', null, pth)); });
      ni.appendChild(ul);
      if (R.length > 1000) ni.appendChild(el('p', 'note', '... and ' + fmtInt(R.length - 1000) + ' more.'));
    }
  }

  // ---------------------------------------------------------------- integrity check
  (function () {
    var sumDirect = 0; for (var d = 0; d < ND; d++) sumDirect += dDirFiles[d];
    var issues = [];
    if (N !== C.files) issues.push('file records (' + N + ') differ from the scan count (' + C.files + ')');
    if (ND > 0 && dTotFiles[0] !== N) issues.push('root total files (' + dTotFiles[0] + ') differ from file records (' + N + ')');
    if (ND > 0 && dTotSize[0] !== totalBytes) issues.push('root total size differs from the sum of file sizes');
    if (sumDirect !== N) issues.push('sum of per-folder file counts differs from file records');
    if (visitOrder.length !== ND) issues.push('some folders are not connected to the tree');
    var f = $('footer');
    f.textContent = 'Generated by ' + META.tool + ' ' + META.version + ' on PowerShell ' + META.psVersion + '. Times are shown in this computer\'s local time zone. ' +
      (issues.length ? 'DATA CHECK FAILED: ' + issues.join('; ') + '.' : 'Data check passed: ' + fmtInt(N) + ' file records, ' + fmtInt(ND) + ' folders, folder totals match file records.');
    if (issues.length) { var b = $('banner'); b.className = 'banner bad'; b.hidden = false; b.textContent = 'Data consistency check failed: ' + issues.join('; ') + '.'; }
  })();

  // ---------------------------------------------------------------- start
  $('app').hidden = false;
  syncDateUi(); syncTypeUi(); renderHead(); renderTree(); refresh();

  // Test hook (read-only): lets automated tests inspect the current view.
  window.__inventory = {
    viewCount: function () { return view.length; },
    viewIndexes: function () { return Array.prototype.slice.call(view); },
    file: function (i) { return { name: fName[i], rel: relPath(i), size: fSize[i], mtime: fM[i], ext: EXT_NAME[fExt[i]], cat: CATS[EXT_CAT[fExt[i]]] }; },
    csv: function () { return buildFilesCsv(view).join(''); },
    state: state, setFolder: setFolder, refresh: refresh, dirIndexByRel: function (rel) { for (var d = 0; d < ND; d++) if (dRel[d] === rel) return d; return -1; },
    dir: function (d) { return { rel: dRel[d], totSize: dTotSize[d], totFiles: dTotFiles[d], dirSize: dDirSize[d], dirFiles: dDirFiles[d], subdirs: dSubdirs[d], descDirs: dDescDirs[d], partial: dPartial[d], status: dStatus[d] }; },
    totalFiles: N, totalBytes: totalBytes, folders: ND
  };
})();
</script>
</body>
</html>
'@

function Write-InventoryReport {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Model,
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][hashtable]$Options
    )
    $token = '__INVENTORY_DATA__'
    $tpl = $script:HtmlTemplate
    $at = $tpl.IndexOf($token)
    if ($at -lt 0 -or $tpl.IndexOf($token, $at + 1) -ge 0) { throw 'Internal error: HTML template data placeholder missing or duplicated.' }
    $head = $tpl.Substring(0, $at)
    $tail = $tpl.Substring($at + $token.Length)

    $now = Get-Date
    $statusText = 'Partial'
    if ($Model.Complete) { $statusText = 'Complete' }
    $meta = [ordered]@{
        tool          = 'FolderInventory'
        version       = $script:ToolVersion
        root          = $Model.RootPath
        sep           = [string][System.IO.Path]::DirectorySeparatorChar
        generated     = $now.ToString('yyyy-MM-dd HH:mm:ss zzz', $script:Invariant)
        generatedUnix = (ConvertTo-UnixSeconds -UtcTime $now.ToUniversalTime())
        durationSec   = [Math]::Round([double]$Model.ElapsedSeconds, 1)
        status        = $statusText
        timedOut      = [bool]$Model.TimedOut
        psVersion     = $PSVersionTable.PSVersion.ToString()
        options       = $Options
        counts        = [ordered]@{
            files = $Model.FileCount; bytes = $Model.TotalBytes; folders = $Model.FolderCount
            skippedFolders = $Model.SkippedFolders; partialFolders = $Model.PartialFolders
            fileErrors = $Model.FileErrors; workerErrors = $Model.WorkerErrors; internalErrors = $Model.InternalErrors
            problems = $Model.Problems.Count; reparseNotFollowed = $Model.Reparse.Count
            excludedFiles = $Model.ExcludedFiles; excludedFolders = $Model.ExcludedDirs
        }
    }

    $statusCode = @{ 'OK' = 0; 'Partial' = 1; 'Skipped' = 2; 'NotScanned' = 3 }
    $rx = $script:JsonEscapeRegex
    $ev = $script:JsonEscapeEvaluator
    $tmp = $Path + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
    $enc = New-Object System.Text.UTF8Encoding -ArgumentList $false
    $writer = New-Object System.IO.StreamWriter -ArgumentList $tmp, $false, $enc, 1048576
    try {
        $writer.Write($head)
        $writer.Write('{"v":1,"meta":')
        $writer.Write((ConvertTo-JsonValue -Value $meta))
        $writer.Write(',"cats":')
        $writer.Write((ConvertTo-JsonValue -Value $script:CategoryNames))

        $writer.Write(',"exts":[')
        for ($e = 0; $e -lt $Model.ExtNames.Count; $e++) {
            if ($e -gt 0) { $writer.Write(',') }
            $writer.Write('[' + (ConvertTo-JsonLiteral -Value $Model.ExtNames[$e]) + ',' + $Model.ExtCategory[$e] + ']')
        }

        # dirs: [parent, name, mtime, status, directSize, directFiles, totalSize, totalFiles,
        #        subdirs, descendantDirs, partial(0/1), attributes, errorType]
        $writer.Write('],"dirs":[')
        $dirs = $Model.Dirs
        for ($i = 0; $i -lt $dirs.Count; $i++) {
            $d = $dirs[$i]
            if ($i -eq 0) { $name = [string]$d.Path } else { $name = [System.IO.Path]::GetFileName([string]$d.Path) }
            if ($rx.IsMatch($name)) { $name = $rx.Replace($name, $ev) }
            $errType = 'null'
            if ($d.ErrType) { $errType = ConvertTo-JsonLiteral -Value ([string]$d.ErrType) }
            $partialFlag = 0
            if ($Model.Partial[$i]) { $partialFlag = 1 }
            $mtime = 0
            if ($null -ne $d.MTime) { $mtime = [long]$d.MTime }
            $attr = 0
            if ($null -ne $d.Attr) { $attr = [int]$d.Attr }
            if ($i -gt 0) { $writer.Write(',') }
            $writer.Write('[' + $Model.Parent[$i] + ',"' + $name + '",' + $mtime + ',' + $statusCode[[string]$d.Status] + ',' +
                $Model.DirectSize[$i] + ',' + $Model.DirectFiles[$i] + ',' + $Model.TotalSize[$i] + ',' + $Model.TotalFiles[$i] + ',' +
                $Model.Subdirs[$i] + ',' + $Model.DescDirs[$i] + ',' + $partialFlag + ',' + $attr + ',' + $errType + ']')
        }

        # files: flat array, 7 values per file:
        #        dirIndex, name, sizeBytes, mtimeUnix, ctimeUnix, attributes, extIndex
        $writer.Write('],"files":[')
        $first = $true
        for ($i = 0; $i -lt $dirs.Count; $i++) {
            $files = $dirs[$i].Files
            if ($null -eq $files) { continue }
            $prefix = [string]$i + ',"'
            foreach ($f in $files) {
                $name = [string]$f[0]
                if ($rx.IsMatch($name)) { $name = $rx.Replace($name, $ev) }
                if ($first) { $first = $false } else { $writer.Write(',') }
                $writer.Write($prefix + $name + '",' + $f[1] + ',' + $f[2] + ',' + $f[3] + ',' + $f[4] + ',' + $f[5])
            }
        }

        $writer.Write('],"problems":[')
        for ($i = 0; $i -lt $Model.Problems.Count; $i++) {
            $pr = $Model.Problems[$i]
            if ($i -gt 0) { $writer.Write(',') }
            $writer.Write('[' + (ConvertTo-JsonLiteral -Value $pr[0]) + ',' + (ConvertTo-JsonLiteral -Value $pr[1]) + ',' +
                (ConvertTo-JsonLiteral -Value $pr[2]) + ',' + (ConvertTo-JsonLiteral -Value $pr[3]) + ']')
        }
        $writer.Write('],"reparse":[')
        for ($i = 0; $i -lt $Model.Reparse.Count; $i++) {
            if ($i -gt 0) { $writer.Write(',') }
            $writer.Write((ConvertTo-JsonLiteral -Value $Model.Reparse[$i]))
        }
        $writer.Write(']}')
        $writer.Write($tail)
    } catch {
        $writer.Dispose()
        if ([System.IO.File]::Exists($tmp)) { [System.IO.File]::Delete($tmp) }
        throw
    }
    $writer.Dispose()

    if ([System.IO.File]::Exists($Path)) { [System.IO.File]::Delete($Path) }
    [System.IO.File]::Move($tmp, $Path)
}

# ---------------------------------------------------------------------------
# Console summary
# ---------------------------------------------------------------------------
function Write-ScanSummary {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Model,
        [Parameter(Mandatory = $true)][string]$ReportPath
    )
    Write-Host ''
    Write-Host '==================== FolderInventory summary ====================' -ForegroundColor Cyan
    Write-Host ('Root:                  {0}' -f $Model.RootPath)
    Write-Host ('Files found:           {0:N0}' -f $Model.FileCount)
    Write-Host ('Folders processed:     {0:N0} (including the root)' -f $Model.FolderCount)
    Write-Host ('Total size:            {0} ({1:N0} bytes)' -f (Format-ByteSize -Bytes $Model.TotalBytes), $Model.TotalBytes)
    Write-Host ('Inaccessible folders:  {0:N0}' -f $Model.SkippedFolders)
    Write-Host ('Partially listed:      {0:N0}' -f $Model.PartialFolders)
    Write-Host ('Errors (all kinds):    {0:N0}' -f $Model.Problems.Count)
    if ($Model.ExcludedFiles -gt 0 -or $Model.ExcludedDirs -gt 0) {
        Write-Host ('Excluded by options:   {0:N0} hidden/system files, {1:N0} hidden/system folders (use -IncludeHidden / -IncludeSystem)' -f $Model.ExcludedFiles, $Model.ExcludedDirs)
    }
    if ($Model.Reparse.Count -gt 0) {
        Write-Host ('Links not followed:    {0:N0} junctions/symbolic links (use -FollowReparsePoints)' -f $Model.Reparse.Count)
    }
    Write-Host ('Scan duration:         {0:N1} s' -f $Model.ElapsedSeconds)
    if ($Model.Complete) {
        Write-Host 'Scan status:           COMPLETE' -ForegroundColor Green
    } else {
        Write-Host 'Scan status:           PARTIAL - some items could not be read; totals are lower bounds.' -ForegroundColor Yellow
        $shown = 0
        foreach ($pr in $Model.Problems) {
            if ($shown -ge 10) { Write-Host ('  ... and {0:N0} more (see the "Skipped & errors" tab in the report)' -f ($Model.Problems.Count - 10)); break }
            Write-Host ('  [{0}] {1}: {2}' -f $pr[1], $pr[3], $pr[0]) -ForegroundColor Yellow
            $shown++
        }
    }
    Write-Host ('Report:                {0}' -f $ReportPath) -ForegroundColor Cyan
    Write-Host ('Open it with:          Start-Process "{0}"' -f $ReportPath)
    Write-Host '=================================================================' -ForegroundColor Cyan
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
$exitCode = 0
try {
    $root = Resolve-InventoryRoot -Path $RootPath
    $outputPath = Resolve-OutputLocation -Path $OutputHtml

    Write-Host ('FolderInventory {0} (PowerShell {1})' -f $script:ToolVersion, $PSVersionTable.PSVersion) -ForegroundColor Cyan
    Write-Host ('Scanning:  {0}' -f $root.Path)
    Write-Host ('Report:    {0}' -f $outputPath)
    Write-Host ('Workers:   {0} | Hidden: {1} | System: {2} | Follow links: {3} | Time limit: {4}' -f $ThrottleLimit, [bool]$IncludeHidden, [bool]$IncludeSystem, [bool]$FollowReparsePoints, $(if ($TimeoutMinutes -gt 0) { "$TimeoutMinutes min" } else { 'none' }))

    $scan = Invoke-FolderScan -Root $root -Workers $ThrottleLimit -Hidden ([bool]$IncludeHidden) -System ([bool]$IncludeSystem) -FollowLinks ([bool]$FollowReparsePoints) -TimeLimitMinutes $TimeoutMinutes
    Write-Host ('Scan finished: {0:N0} folders, {1:N0} files. Building report...' -f $scan.Dirs.Count, $scan.FileCount)

    $model = Build-InventoryModel -Scan $scan -Root $root
    $options = [ordered]@{
        throttleLimit = $ThrottleLimit; includeHidden = [bool]$IncludeHidden; includeSystem = [bool]$IncludeSystem
        followReparsePoints = [bool]$FollowReparsePoints; timeoutMinutes = $TimeoutMinutes; maxDepth = $script:MaxDepth
    }
    Write-InventoryReport -Model $model -Path $outputPath -Options $options
    Write-ScanSummary -Model $model -ReportPath $outputPath
    if (-not $model.Complete) { $exitCode = 2 }
} catch {
    $exitCode = 1
    $err = $_
    $isUserError = $false
    try { $isUserError = [bool]$err.Exception.Data['FolderInventoryUserError'] } catch { $isUserError = $false }
    Write-Host ''
    Write-Host ('FATAL: {0}' -f $err.Exception.Message) -ForegroundColor Red
    if (-not $isUserError) {
        Write-Host ('  Exception: {0}' -f $err.Exception.GetType().FullName) -ForegroundColor Red
        if ($err.InvocationInfo) {
            Write-Host ('  At line {0}: {1}' -f $err.InvocationInfo.ScriptLineNumber, ([string]$err.InvocationInfo.Line).Trim()) -ForegroundColor Red
        }
        if ($err.ScriptStackTrace) { Write-Host ('  Stack:' + [Environment]::NewLine + $err.ScriptStackTrace) -ForegroundColor DarkRed }
    }
    Write-Host 'No report was written.' -ForegroundColor Red
}
exit $exitCode
