<#
.SYNOPSIS
    Shared functions library for the file organisation suite.
    Dot-sourced by all scripts — do not run this file directly.

    Scripts that use this library:
      Invoke-Dedup.ps1
      Invoke-SortPhotos.ps1
      Invoke-SortVideos.ps1
      Invoke-SortDocuments.ps1  (and any future scripts)

.FUNCTIONS
    File Transfer
      Invoke-FileTransfer       — copy/move a file with SHA256 collision detection
      Compress-OutputFolder     — zip a folder into a timestamped archive

    Hashing
      Get-PartialHash           — fast MD5 of first+last 64KB (pre-filter)
      Get-FullHash              — full SHA256 of a file (definitive match)

    Date Resolution
      Get-LastWriteTimeFallback — guaranteed LastWriteTime fallback for any file
      Get-MonthFolderName       — formats a date into a yyyy\month folder path

    File Collection
      Get-FilteredFiles         — scans a folder with optional extension filter + early-exit

    Validation
      Assert-SourcePath         — validates SourcePath exists and is accessible
#>

Set-StrictMode -Version Latest

# =============================================================================
# FILE TRANSFER
# =============================================================================

function Invoke-FileTransfer {
    <#
    .SYNOPSIS
        Copies or moves a file to a destination folder with robust collision handling.

    .DESCRIPTION
        Collision resolution uses a three-stage check — much safer than the old
        size-only approach which could silently skip genuinely different files:

          Stage 1 — Different file size?
                    → Rename incoming with _1, _2 ... suffix (definitely different)

          Stage 2 — Same size, same SHA256?
                    → Skip (confirmed identical content — true duplicate)

          Stage 3 — Same size, different SHA256?
                    → Rename incoming (different content, coincidental size match)

        This guarantees no file is ever silently lost due to a false size match.

    .PARAMETER FilePath
        Full path to the source file.

    .PARAMETER DestFolder
        Destination folder. Created automatically if it does not exist.

    .PARAMETER DeleteSource
        If $true the source file is deleted after a successful copy (Move mode).
        The copy always completes before any delete — no data loss on failure.

    .OUTPUTS
        PSCustomObject with:
          Path   — full path of the destination file
          Action — 'Copied' | 'Moved' | 'Skipped (identical)' | 'Skipped (WhatIf)'
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)][string] $FilePath,
        [Parameter(Mandatory)][string] $DestFolder,
        [bool] $DeleteSource = $false
    )

    # ── Ensure destination folder exists ─────────────────────────────────────
    if (-not (Test-Path $DestFolder)) {
        if ($PSCmdlet.ShouldProcess($DestFolder, 'Create directory')) {
            New-Item -ItemType Directory -Path $DestFolder -Force | Out-Null
        }
    }

    $fileName  = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    $extension = [System.IO.Path]::GetExtension($FilePath)
    $destPath  = Join-Path $DestFolder "$fileName$extension"

    # ── Collision resolution ──────────────────────────────────────────────────
    $counter  = 1
    while (Test-Path $destPath) {
        $existing = Get-Item -LiteralPath $destPath
        $incoming = Get-Item -LiteralPath $FilePath

        if ($existing.Length -ne $incoming.Length) {
            # Stage 1 — Different size: definitely different files, rename
            $destPath = Join-Path $DestFolder "${fileName}_$counter$extension"
            $counter++
            continue
        }

        # Stage 2 & 3 — Same size: SHA256 decides
        $hashExisting = Get-FullHash -FilePath $existing.FullName
        $hashIncoming = Get-FullHash -FilePath $FilePath

        if ($hashExisting -eq $hashIncoming) {
            # Stage 2 — Identical content: skip, it's already there
            Write-Verbose "Skipped (identical): '$FilePath' already exists at '$destPath'"
            return [PSCustomObject]@{ Path = $destPath; Action = 'Skipped (identical)' }
        }
        else {
            # Stage 3 — Same size, different content: rename incoming
            $destPath = Join-Path $DestFolder "${fileName}_$counter$extension"
            $counter++
        }
    }

    # ── Transfer ──────────────────────────────────────────────────────────────
    $action = if ($DeleteSource) { 'Move' } else { 'Copy' }

    if ($PSCmdlet.ShouldProcess($destPath, "$action '$FilePath'")) {
        Copy-Item -LiteralPath $FilePath -Destination $destPath -Force

        if ($DeleteSource) {
            Remove-Item -LiteralPath $FilePath -Force
        }

        return [PSCustomObject]@{ Path = $destPath; Action = if ($DeleteSource) { 'Moved' } else { 'Copied' } }
    }

    # WhatIf path — no file was touched
    return [PSCustomObject]@{ Path = $destPath; Action = 'Skipped (WhatIf)' }
}

# =============================================================================
# COMPRESSION
# =============================================================================

function Compress-OutputFolder {
    <#
    .SYNOPSIS
        Compresses a folder into a timestamped .zip archive using built-in
        System.IO.Compression — no third-party tools required.

    .PARAMETER FolderPath
        The folder to compress.

    .PARAMETER ZipDestFolder
        Where to place the resulting .zip file.

    .PARAMETER Label
        Short label used in the zip filename, e.g. 'source' or 'copies'.
        Output filename: label_yyyyMMddTHHmmssZ.zip

    .OUTPUTS
        Full path to the created zip file.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)][string] $FolderPath,
        [Parameter(Mandatory)][string] $ZipDestFolder,
        [Parameter(Mandatory)][string] $Label
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $timestamp = (Get-Date).ToUniversalTime().ToString('yyyyMMddTHHmmssZ')
    $zipName   = "${Label}_${timestamp}.zip"
    $zipPath   = Join-Path $ZipDestFolder $zipName

    Write-Host "  Compressing '$FolderPath'" -ForegroundColor Cyan
    Write-Host "           → '$zipPath'..."  -ForegroundColor Cyan

    if ($PSCmdlet.ShouldProcess($zipPath, "Create zip archive of '$FolderPath'")) {
        [System.IO.Compression.ZipFile]::CreateFromDirectory(
            $FolderPath,
            $zipPath,
            [System.IO.Compression.CompressionLevel]::Optimal,
            $false   # do not include the base directory name inside the zip
        )
        $sizeMB = [math]::Round((Get-Item $zipPath).Length / 1MB, 2)
        Write-Host "  Done. Archive: ${sizeMB} MB" -ForegroundColor Green
    }

    return $zipPath
}

# =============================================================================
# HASHING
# =============================================================================

function Get-PartialHash {
    <#
    .SYNOPSIS
        Fast MD5 hash of the first and last 64 KB of a file.
        Used as a cheap pre-filter before committing to a full SHA256.
        Falls back to full content for files smaller than 128 KB.
    #>
    param ([Parameter(Mandatory)][string]$FilePath)

    $chunkSize = 64KB
    $bytes     = [System.Collections.Generic.List[byte]]::new()
    $stream    = [System.IO.File]::OpenRead($FilePath)

    try {
        $len = $stream.Length

        if ($len -le ($chunkSize * 2)) {
            # Small file — read everything
            $buf = [byte[]]::new($len)
            $null = $stream.Read($buf, 0, $len)
            $bytes.AddRange($buf)
        }
        else {
            # First 64 KB
            $buf = [byte[]]::new($chunkSize)
            $null = $stream.Read($buf, 0, $chunkSize)
            $bytes.AddRange($buf)
            # Last 64 KB
            $null = $stream.Seek(-$chunkSize, [System.IO.SeekOrigin]::End)
            $buf  = [byte[]]::new($chunkSize)
            $null = $stream.Read($buf, 0, $chunkSize)
            $bytes.AddRange($buf)
        }
    }
    finally { $stream.Dispose() }

    $md5 = [System.Security.Cryptography.MD5]::Create()
    try {
        return [System.BitConverter]::ToString($md5.ComputeHash($bytes.ToArray())) -replace '-', ''
    }
    finally { $md5.Dispose() }
}

function Get-FullHash {
    <#
    .SYNOPSIS
        Full SHA256 hash of a file. Used for definitive duplicate confirmation
        and for collision resolution in Invoke-FileTransfer.
    #>
    param ([Parameter(Mandatory)][string]$FilePath)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $stream = [System.IO.File]::OpenRead($FilePath)
        try {
            return [System.BitConverter]::ToString($sha.ComputeHash($stream)) -replace '-', ''
        }
        finally { $stream.Dispose() }
    }
    finally { $sha.Dispose() }
}

# =============================================================================
# DATE RESOLUTION
# =============================================================================

function Get-LastWriteTimeFallback {
    <#
    .SYNOPSIS
        Returns a file's LastWriteTime as a guaranteed date fallback.
        Used by all sort scripts when embedded metadata is absent.

    .DESCRIPTION
        LastWriteTime is preferred over CreationTime because Windows resets
        CreationTime every time a file is copied, whereas LastWriteTime is
        preserved across most copy operations.

        No year-range gate is applied — any valid timestamp is accepted so
        that no file is ever discarded simply due to an unusual date.

    .OUTPUTS
        PSCustomObject with Date ([datetime]) and Source ([string]),
        or $null if the file system cannot provide a valid date (rare).
    #>
    param ([Parameter(Mandatory)][System.IO.FileInfo]$File)

    $lwt = $File.LastWriteTime
    if ($null -ne $lwt -and $lwt -ne [datetime]::MinValue) {
        return [PSCustomObject]@{
            Date   = $lwt
            Source = 'LastWriteTime (metadata absent — fallback)'
        }
    }
    return $null
}

function Get-MonthFolderName {
    <#
    .SYNOPSIS
        Converts a [datetime] into a month subfolder name string.

    .PARAMETER Date
        The date to format.

    .PARAMETER Format
        Named   →  06 - June     (default, human-readable)
        Number  →  06            (numeric, compact)
        Full    →  2024-06       (ISO-style year+month)
    #>
    param (
        [Parameter(Mandatory)][datetime] $Date,
        [ValidateSet('Named', 'Number', 'Full')]
        [string] $Format = 'Named'
    )

    $monthNames = @{
        1='January'; 2='February'; 3='March';     4='April'
        5='May';     6='June';     7='July';       8='August'
        9='September'; 10='October'; 11='November'; 12='December'
    }

    switch ($Format) {
        'Number' { return '{0:D2}'         -f $Date.Month }
        'Full'   { return '{0}-{1:D2}'     -f $Date.Year, $Date.Month }
        default  { return '{0:D2} - {1}'   -f $Date.Month, $monthNames[$Date.Month] }
    }
}

# =============================================================================
# FILE COLLECTION
# =============================================================================

function Get-FilteredFiles {
    <#
    .SYNOPSIS
        Scans a folder for files, optionally filtered by extension whitelist
        and/or extension blacklist. Returns an array (never $null).

    .DESCRIPTION
        Always returns [array] — wraps results in @() so .Count is always safe
        even when zero files match, eliminating the null-.Count bug that
        previously affected all scripts under Set-StrictMode.

    .PARAMETER FolderPath
        The folder to scan.

    .PARAMETER Recurse
        If specified, scans subfolders too.

    .PARAMETER IncludeExtensions
        If provided, only files with these extensions are returned.
        Extensions must include the leading dot, e.g. '.jpg', '.pdf'

    .PARAMETER ExcludeExtensions
        If provided, files with these extensions are excluded.
        Applied after IncludeExtensions. Extensions must include the leading dot.

    .PARAMETER EmptyMessage
        Custom message shown when no files are found (optional).

    .OUTPUTS
        [array] of System.IO.FileInfo objects. Never $null.
    #>
    param (
        [Parameter(Mandatory)][string]   $FolderPath,
        [switch]                         $Recurse,
        [string[]]                       $IncludeExtensions = @(),
        [string[]]                       $ExcludeExtensions = @(),
        [string]                         $EmptyMessage      = ''
    )

    $params = @{ LiteralPath = $FolderPath; File = $true }
    if ($Recurse) { $params['Recurse'] = $true }

    $files = @(Get-ChildItem @params | Where-Object {
        $ext = $_.Extension.ToLower()

        $included = ($IncludeExtensions.Count -eq 0) -or ($ext -in $IncludeExtensions)
        $excluded = ($ExcludeExtensions.Count -gt 0)  -and ($ext -in $ExcludeExtensions)

        $included -and -not $excluded
    })

    if ($files.Count -eq 0) {
        $msg = if ($EmptyMessage) { $EmptyMessage } else {
            "  No matching files found in '$FolderPath'."
        }
        Write-Host $msg -ForegroundColor Yellow
    }

    return $files
}

# =============================================================================
# VALIDATION
# =============================================================================

function Assert-SourcePath {
    <#
    .SYNOPSIS
        Validates that SourcePath exists and is a folder.
        Throws a clear error if not — prevents cryptic downstream failures.
    #>
    param ([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "SourcePath not found: '$Path'"
    }
    if (-not (Get-Item -LiteralPath $Path).PSIsContainer) {
        throw "SourcePath is not a folder: '$Path'"
    }
}
