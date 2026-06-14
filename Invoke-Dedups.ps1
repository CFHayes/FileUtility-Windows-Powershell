<#
.SYNOPSIS
    Deduplicates files using a tiered approach:
    1. Group by file size + extension (fast metadata filter)
    2. Partial hash of first/last 64KB (cheap secondary filter)
    3. Full SHA256 hash (confirmed duplicate detection)

.DESCRIPTION
    Scans a source folder and sorts files into:
      - originals\  : unique files or the first-seen copy of a duplicate
      - copies\     : confirmed duplicate files

    By default files are COPIED and your source is left untouched.
    Use the switches below to change post-processing behaviour.

.PARAMETER SourcePath
    Path to the folder containing files to deduplicate.

.PARAMETER OutputPath
    Path to the output folder. 'originals' and 'copies' subfolders
    will be created here automatically.

.PARAMETER Recurse
    Scan subfolders of SourcePath as well.

.PARAMETER Move
    Move files instead of copying them. The source file is removed after
    it has been successfully transferred. Cannot be combined with -DeleteCopies.

.PARAMETER DeleteCopies
    After duplicates are copied to copies\, delete them from the source folder.
    Originals in the source folder are left untouched.
    Cannot be combined with -Move.

.PARAMETER Compress
    Compress one or both targets into a .zip archive after processing.
    Valid values:
      Source   – zip the original SourcePath folder
      Copies   – zip the copies\ output folder
      Both     – zip both

    The .zip file is placed in OutputPath and named with a UTC timestamp,
    e.g. source_20240615T123456Z.zip / copies_20240615T123456Z.zip

.PARAMETER WhatIf
    Dry run — reports what would happen without touching any files.

.EXAMPLE
    # Basic dedup — copy everything, touch nothing in source
    .\Invoke-Dedup.ps1 -SourcePath "C:\MyFiles" -OutputPath "C:\Sorted"

.EXAMPLE
    # Move files out of source instead of copying
    .\Invoke-Dedup.ps1 -SourcePath "C:\MyFiles" -OutputPath "C:\Sorted" -Move

.EXAMPLE
    # Copy, then delete duplicates from source; zip the copies\ folder
    .\Invoke-Dedup.ps1 -SourcePath "C:\MyFiles" -OutputPath "C:\Sorted" -DeleteCopies -Compress Copies

.EXAMPLE
    # Recurse, move, compress both source and copies\, dry run first
    .\Invoke-Dedup.ps1 -SourcePath "C:\MyFiles" -OutputPath "C:\Sorted" -Recurse -Move -Compress Both -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory)][string] $SourcePath,
    [Parameter(Mandatory)][string] $OutputPath,
    [switch] $Recurse,

    # Mutually exclusive transfer modes
    [switch] $Move,
    [switch] $DeleteCopies,

    # Compression target
    [ValidateSet('Source', 'Copies', 'Both')]
    [string] $Compress
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Parameter validation
# ---------------------------------------------------------------------------

if ($Move -and $DeleteCopies) {
    throw "Cannot use -Move and -DeleteCopies together. " +
          "-Move transfers ALL files; -DeleteCopies only removes duplicates from source."
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Get-PartialHash {
    <#
    Reads the first and last 64 KB of a file and returns their combined MD5.
    Falls back to full content for files smaller than 128 KB.
    #>
    param ([string]$FilePath)

    $chunkSize = 64KB
    $bytes = [System.Collections.Generic.List[byte]]::new()

    $stream = [System.IO.File]::OpenRead($FilePath)
    try {
        $fileLength = $stream.Length

        if ($fileLength -le ($chunkSize * 2)) {
            $buf = [byte[]]::new($fileLength)
            $null = $stream.Read($buf, 0, $fileLength)
            $bytes.AddRange($buf)
        }
        else {
            # First 64 KB
            $buf = [byte[]]::new($chunkSize)
            $null = $stream.Read($buf, 0, $chunkSize)
            $bytes.AddRange($buf)

            # Last 64 KB
            $null = $stream.Seek(-$chunkSize, [System.IO.SeekOrigin]::End)
            $buf = [byte[]]::new($chunkSize)
            $null = $stream.Read($buf, 0, $chunkSize)
            $bytes.AddRange($buf)
        }
    }
    finally { $stream.Dispose() }

    $md5 = [System.Security.Cryptography.MD5]::Create()
    try {
        $hash = $md5.ComputeHash($bytes.ToArray())
        return [System.BitConverter]::ToString($hash) -replace '-', ''
    }
    finally { $md5.Dispose() }
}

function Get-FullHash {
    param ([string]$FilePath)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $stream = [System.IO.File]::OpenRead($FilePath)
        try {
            $hash = $sha.ComputeHash($stream)
            return [System.BitConverter]::ToString($hash) -replace '-', ''
        }
        finally { $stream.Dispose() }
    }
    finally { $sha.Dispose() }
}

function Move-ToOutput {
    <#
    Copies a file to $DestFolder (handling name collisions), then — if not a
    WhatIf run — deletes the source file.  Returns a result object with Path and Action.
    #>
    param (
        [string] $FilePath,
        [string] $DestFolder,
        [bool]   $DeleteSource   # true = Move mode or DeleteCopies for duplicate
    )

    # Create destination folder if it doesn't exist yet
    if (-not (Test-Path $DestFolder)) {
        if ($PSCmdlet.ShouldProcess($DestFolder, 'Create directory')) {
            New-Item -ItemType Directory -Path $DestFolder -Force | Out-Null
        }
    }

    $fileName  = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    $extension = [System.IO.Path]::GetExtension($FilePath)
    $destPath  = Join-Path $DestFolder "$fileName$extension"

    # Collision handling — skip if same name + size (already processed),
    # rename with incrementing suffix otherwise
    $counter = 1
    while (Test-Path $destPath) {
        $existing = Get-Item $destPath
        $incoming = Get-Item $FilePath
        if ($existing.Length -eq $incoming.Length) {
            return [PSCustomObject]@{ Path = $destPath; Action = 'Skipped (already exists)' }
        }
        $destPath = Join-Path $DestFolder "${fileName}_$counter$extension"
        $counter++
    }

    if ($PSCmdlet.ShouldProcess($destPath, "Copy '$FilePath'")) {
        Copy-Item -LiteralPath $FilePath -Destination $destPath -Force
    }

    if ($DeleteSource -and $PSCmdlet.ShouldProcess($FilePath, 'Delete source file')) {
        Remove-Item -LiteralPath $FilePath -Force
    }

    return [PSCustomObject]@{ Path = $destPath; Action = if ($DeleteSource) { 'Moved' } else { 'Copied' } }
}

function Compress-Folder {
    <#
    Compresses $FolderPath into a timestamped .zip in $ZipDestFolder.
    Uses the built-in System.IO.Compression — no external tools needed.
    #>
    param (
        [string] $FolderPath,
        [string] $ZipDestFolder,
        [string] $Label           # 'source' or 'copies'
    )

    $timestamp = (Get-Date).ToUniversalTime().ToString('yyyyMMddTHHmmssZ')
    $zipName   = "${Label}_${timestamp}.zip"
    $zipPath   = Join-Path $ZipDestFolder $zipName

    Write-Host "  Compressing '$FolderPath' → '$zipPath'..." -ForegroundColor Cyan

    if ($PSCmdlet.ShouldProcess($zipPath, "Create zip archive of '$FolderPath'")) {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::CreateFromDirectory(
            $FolderPath,
            $zipPath,
            [System.IO.Compression.CompressionLevel]::Optimal,
            $false   # don't include base directory name inside zip
        )
        $zipSizeMB = [math]::Round((Get-Item $zipPath).Length / 1MB, 2)
        Write-Host "  Done. Archive size: ${zipSizeMB} MB  →  $zipPath" -ForegroundColor Green
    }

    return $zipPath
}

# ---------------------------------------------------------------------------
# Setup output folders
# ---------------------------------------------------------------------------

$originalsFolder = Join-Path $OutputPath 'originals'
$copiesFolder    = Join-Path $OutputPath 'copies'

foreach ($folder in @($originalsFolder, $copiesFolder)) {
    if (-not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder | Out-Null }
}

# Describe the run mode clearly upfront
$modeDescription = switch ($true) {
    ($Move)         { 'MOVE  (source files will be removed after transfer)' }
    ($DeleteCopies) { 'COPY + DELETE DUPLICATES FROM SOURCE' }
    default         { 'COPY  (source files untouched)' }
}
Write-Host "`nMode : $modeDescription" -ForegroundColor Yellow
if ($Compress) {
    Write-Host "Compress target : $Compress" -ForegroundColor Yellow
}

# ---------------------------------------------------------------------------
# Step 1 — Collect files
# ---------------------------------------------------------------------------

Write-Host "`n[1/4] Scanning '$SourcePath'..." -ForegroundColor Cyan

$getChildParams = @{ LiteralPath = $SourcePath; File = $true }
if ($Recurse) { $getChildParams['Recurse'] = $true }

# @() guarantees an array even when the folder is empty or returns a single item
$allFiles = @(Get-ChildItem @getChildParams)

if ($allFiles.Count -eq 0) {
    Write-Host "`n  No files found in '$SourcePath'." -ForegroundColor Yellow
    exit 0
}

Write-Host "      Found $($allFiles.Count) file(s)."

# ---------------------------------------------------------------------------
# Step 2 — Group by size + extension (metadata pre-filter)
# ---------------------------------------------------------------------------

Write-Host "`n[2/4] Grouping by size + extension..." -ForegroundColor Cyan

$groups = $allFiles | Group-Object -Property {
    "$($_.Length)|$($_.Extension.ToLower())"
}

# @() prevents null when all files are unique (no candidates) or all are dupes (no singletons)
$singletons      = @($groups | Where-Object { $_.Count -eq 1 })
$candidateGroups = @($groups | Where-Object { $_.Count -gt 1 })

# Measure-Object returns an object with Sum=$null when the collection is empty — default to 0
$singletonCount = if ($singletons.Count      -gt 0) { ($singletons      | Measure-Object -Property Count -Sum).Sum } else { 0 }
$candidateCount = if ($candidateGroups.Count -gt 0) { ($candidateGroups | Measure-Object -Property Count -Sum).Sum } else { 0 }

Write-Host "      Singletons (unique size+ext): $singletonCount file(s) → originals/"
Write-Host "      Candidates for hashing:       $candidateCount file(s)"

# Initialise stats before the singleton loop so Skipped can be incremented there
$stats = @{ Originals = [int]$singletonCount; Copies = 0; Errors = 0; Deleted = 0; Skipped = 0 }

# Singletons can never be duplicates — send straight to originals
# In Move mode we delete the source; in Copy/DeleteCopies mode we leave it.
foreach ($group in $singletons) {
    $file   = $group.Group[0]
    $result = Move-ToOutput -FilePath $file.FullName -DestFolder $originalsFolder -DeleteSource $Move.IsPresent
    if ($result.Action -eq 'Skipped (already exists)') { $stats.Skipped++ }
}

# ---------------------------------------------------------------------------
# Step 3 — Partial hash on candidates
# ---------------------------------------------------------------------------

Write-Host "`n[3/4] Running partial hash on candidates..." -ForegroundColor Cyan

$partialHashGroups = [System.Collections.Generic.Dictionary[string,
    System.Collections.Generic.List[string]]]::new()

$candidateFiles = $candidateGroups | ForEach-Object { $_.Group }
$total  = @($candidateFiles).Count
$index  = 0

foreach ($file in $candidateFiles) {
    $index++
    Write-Progress -Activity 'Partial hashing' `
                   -Status   "$index / $total : $($file.Name)" `
                   -PercentComplete (($index / $total) * 100)
    try {
        $ph = Get-PartialHash -FilePath $file.FullName
        if (-not $partialHashGroups.ContainsKey($ph)) {
            $partialHashGroups[$ph] = [System.Collections.Generic.List[string]]::new()
        }
        $partialHashGroups[$ph].Add($file.FullName)
    }
    catch {
        Write-Warning "Could not partial-hash '$($file.FullName)': $_"
        $stats.Errors++
    }
}
Write-Progress -Activity 'Partial hashing' -Completed

# Partial-hash singletons are unique — no need for full hash
foreach ($entry in @($partialHashGroups.GetEnumerator() | Where-Object { $_.Value.Count -eq 1 })) {
    $result = Move-ToOutput -FilePath $entry.Value[0] -DestFolder $originalsFolder -DeleteSource $Move.IsPresent
    if ($result.Action -eq 'Skipped (already exists)') { $stats.Skipped++ }
    else { $stats.Originals++ }
}

# ---------------------------------------------------------------------------
# Step 4 — Full SHA256 on partial-hash matches
# ---------------------------------------------------------------------------

Write-Host "`n[4/4] Running full SHA256 on partial-hash matches..." -ForegroundColor Cyan

$seenHashes  = @{}   # SHA256 → first-seen source path
$phGroupList = @($partialHashGroups.GetEnumerator() | Where-Object { $_.Value.Count -gt 1 })
# Default to 0 — Measure-Object returns Sum=$null on an empty collection
$total  = if ($phGroupList.Count -gt 0) {
    ($phGroupList | ForEach-Object { $_.Value.Count } | Measure-Object -Sum).Sum
} else { 0 }
$index  = 0

if ($total -eq 0) {
    Write-Host "      No full-hash candidates — all duplicates resolved by partial hash." -ForegroundColor Cyan
}

foreach ($entry in $phGroupList) {
    foreach ($filePath in $entry.Value) {
        $index++
        Write-Progress -Activity 'Full SHA256' `
                       -Status   "$index / $total : $(Split-Path $filePath -Leaf)" `
                       -PercentComplete (($index / $total) * 100)
        try {
            $fullHash = Get-FullHash -FilePath $filePath

            if ($seenHashes.ContainsKey($fullHash)) {
                # ── DUPLICATE ──────────────────────────────────────────────
                # Always copy to copies\ first so we have a record
                Move-ToOutput -FilePath $filePath `
                              -DestFolder $copiesFolder `
                              -DeleteSource ($Move -or $DeleteCopies) | Out-Null
                $stats.Copies++
                if ($Move -or $DeleteCopies) { $stats.Deleted++ }
            }
            else {
                # ── ORIGINAL (first seen) ──────────────────────────────────
                $seenHashes[$fullHash] = $filePath
                Move-ToOutput -FilePath $filePath `
                              -DestFolder $originalsFolder `
                              -DeleteSource $Move.IsPresent | Out-Null
                $stats.Originals++
            }
        }
        catch {
            Write-Warning "Could not full-hash '$filePath': $_"
            $stats.Errors++
        }
    }
}
Write-Progress -Activity 'Full SHA256' -Completed

# ---------------------------------------------------------------------------
# Compression (optional)
# ---------------------------------------------------------------------------

if ($Compress) {
    Write-Host "`n[+] Compressing..." -ForegroundColor Cyan

    if ($Compress -in @('Source', 'Both')) {
        Compress-Folder -FolderPath $SourcePath -ZipDestFolder $OutputPath -Label 'source'
    }
    if ($Compress -in @('Copies', 'Both')) {
        Compress-Folder -FolderPath $copiesFolder -ZipDestFolder $OutputPath -Label 'copies'
    }
}

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "  Deduplication Complete" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Mode                 : $modeDescription"
Write-Host "  Total files scanned  : $($allFiles.Count)"
Write-Host "  Originals            : $($stats.Originals)  →  $originalsFolder"
Write-Host "  Copies (duplicates)  : $($stats.Copies)  →  $copiesFolder"
if ($Move) {
    Write-Host "  Source files removed : $($stats.Deleted)  (Move mode)"  -ForegroundColor Yellow
}
elseif ($DeleteCopies) {
    Write-Host "  Duplicates deleted   : $($stats.Deleted)  from source"  -ForegroundColor Yellow
}
if ($Compress) {
    Write-Host "  Compressed           : $Compress  →  $OutputPath"
}
if ($stats.Skipped -gt 0) {
    Write-Host "  Skipped (already exists) : $($stats.Skipped)" -ForegroundColor Yellow
}
if ($stats.Errors -gt 0) {
    Write-Host "  Errors               : $($stats.Errors)" -ForegroundColor Red
}
Write-Host "========================================`n" -ForegroundColor Green
