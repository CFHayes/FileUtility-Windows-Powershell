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

    Requires Shared-Functions.ps1 in the same folder.

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
    Valid values:  Source | Copies | Both
    The .zip is placed in OutputPath with a UTC timestamp in the filename.

.PARAMETER WhatIf
    Dry run — reports what would happen without touching any files.

.EXAMPLE
    .\Invoke-Dedup.ps1 -SourcePath "C:\MyFiles" -OutputPath "C:\Sorted"

.EXAMPLE
    .\Invoke-Dedup.ps1 -SourcePath "C:\MyFiles" -OutputPath "C:\Sorted" -Move

.EXAMPLE
    .\Invoke-Dedup.ps1 -SourcePath "C:\MyFiles" -OutputPath "C:\Sorted" -DeleteCopies -Compress Copies

.EXAMPLE
    .\Invoke-Dedup.ps1 -SourcePath "C:\MyFiles" -OutputPath "C:\Sorted" -Recurse -Move -Compress Both -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory)][string] $SourcePath,
    [Parameter(Mandatory)][string] $OutputPath,
    [switch] $Recurse,
    [switch] $Move,
    [switch] $DeleteCopies,
    [ValidateSet('Source', 'Copies', 'Both')]
    [string] $Compress
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Load shared functions
# ---------------------------------------------------------------------------

. "$PSScriptRoot\Shared-Functions.ps1"

# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------

Assert-SourcePath -Path $SourcePath

if ($Move -and $DeleteCopies) {
    throw "Cannot use -Move and -DeleteCopies together. " +
          "-Move transfers ALL files; -DeleteCopies only removes duplicates from source."
}

# ---------------------------------------------------------------------------
# Setup
# ---------------------------------------------------------------------------

$originalsFolder = Join-Path $OutputPath 'originals'
$copiesFolder    = Join-Path $OutputPath 'copies'

foreach ($folder in @($originalsFolder, $copiesFolder)) {
    if (-not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder | Out-Null }
}

$modeDescription = switch ($true) {
    ($Move)         { 'MOVE  (source files will be removed after transfer)' }
    ($DeleteCopies) { 'COPY + DELETE DUPLICATES FROM SOURCE' }
    default         { 'COPY  (source files untouched)' }
}

Write-Host "`nMode : $modeDescription" -ForegroundColor Yellow
if ($Compress) { Write-Host "Compress target : $Compress" -ForegroundColor Yellow }

# ---------------------------------------------------------------------------
# Step 1 — Collect files
# ---------------------------------------------------------------------------

Write-Host "`n[1/4] Scanning '$SourcePath'..." -ForegroundColor Cyan

$allFiles = Get-FilteredFiles -FolderPath $SourcePath -Recurse:$Recurse `
                              -EmptyMessage "  No files found in '$SourcePath'."

if ($allFiles.Count -eq 0) { exit 0 }

Write-Host "      Found $($allFiles.Count) file(s)."

# ---------------------------------------------------------------------------
# Step 2 — Group by size + extension (metadata pre-filter)
# ---------------------------------------------------------------------------

Write-Host "`n[2/4] Grouping by size + extension..." -ForegroundColor Cyan

$groups          = $allFiles | Group-Object -Property { "$($_.Length)|$($_.Extension.ToLower())" }
$singletons      = @($groups | Where-Object { $_.Count -eq 1 })
$candidateGroups = @($groups | Where-Object { $_.Count -gt 1 })

$singletonCount = if ($singletons.Count      -gt 0) { ($singletons      | Measure-Object -Property Count -Sum).Sum } else { 0 }
$candidateCount = if ($candidateGroups.Count -gt 0) { ($candidateGroups | Measure-Object -Property Count -Sum).Sum } else { 0 }

Write-Host "      Singletons (unique size+ext): $singletonCount file(s) → originals/"
Write-Host "      Candidates for hashing:       $candidateCount file(s)"

$stats = @{ Originals = [int]$singletonCount; Copies = 0; Errors = 0; Deleted = 0; Skipped = 0 }

# Singletons cannot be duplicates — send straight to originals
foreach ($group in $singletons) {
    $result = Invoke-FileTransfer -FilePath     $group.Group[0].FullName `
                                  -DestFolder   $originalsFolder `
                                  -DeleteSource $Move.IsPresent
    if ($result.Action -like 'Skipped*') { $stats.Skipped++ }
}

# ---------------------------------------------------------------------------
# Step 3 — Partial hash on candidates
# ---------------------------------------------------------------------------

Write-Host "`n[3/4] Running partial hash on candidates..." -ForegroundColor Cyan

$partialHashGroups = [System.Collections.Generic.Dictionary[string,
    System.Collections.Generic.List[string]]]::new()

$candidateFiles = @($candidateGroups | ForEach-Object { $_.Group })
$total  = $candidateFiles.Count
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

# Partial-hash singletons are unique — no full hash needed
foreach ($entry in @($partialHashGroups.GetEnumerator() | Where-Object { $_.Value.Count -eq 1 })) {
    $result = Invoke-FileTransfer -FilePath     $entry.Value[0] `
                                  -DestFolder   $originalsFolder `
                                  -DeleteSource $Move.IsPresent
    if ($result.Action -like 'Skipped*') { $stats.Skipped++ }
    else { $stats.Originals++ }
}

# ---------------------------------------------------------------------------
# Step 4 — Full SHA256 on partial-hash matches
# ---------------------------------------------------------------------------

Write-Host "`n[4/4] Running full SHA256 on partial-hash matches..." -ForegroundColor Cyan

$seenHashes  = @{}
$phGroupList = @($partialHashGroups.GetEnumerator() | Where-Object { $_.Value.Count -gt 1 })
$total = if ($phGroupList.Count -gt 0) {
    ($phGroupList | ForEach-Object { $_.Value.Count } | Measure-Object -Sum).Sum
} else { 0 }
$index = 0

if ($total -eq 0) {
    Write-Host "      No full-hash candidates — all resolved by partial hash." -ForegroundColor Cyan
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
                $result = Invoke-FileTransfer -FilePath     $filePath `
                                              -DestFolder   $copiesFolder `
                                              -DeleteSource ($Move -or $DeleteCopies)
                $stats.Copies++
                if ($Move -or $DeleteCopies) { $stats.Deleted++ }
            }
            else {
                $seenHashes[$fullHash] = $filePath
                $result = Invoke-FileTransfer -FilePath     $filePath `
                                              -DestFolder   $originalsFolder `
                                              -DeleteSource $Move.IsPresent
                if ($result.Action -like 'Skipped*') { $stats.Skipped++ }
                else { $stats.Originals++ }
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
# Compression
# ---------------------------------------------------------------------------

if ($Compress) {
    Write-Host "`n[+] Compressing..." -ForegroundColor Cyan
    if ($Compress -in @('Source', 'Both')) {
        Compress-OutputFolder -FolderPath $SourcePath    -ZipDestFolder $OutputPath -Label 'source'
    }
    if ($Compress -in @('Copies', 'Both')) {
        Compress-OutputFolder -FolderPath $copiesFolder  -ZipDestFolder $OutputPath -Label 'copies'
    }
}

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "  Deduplication Complete"                  -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Mode                 : $modeDescription"
Write-Host "  Total files scanned  : $($allFiles.Count)"
Write-Host "  Originals            : $($stats.Originals)  →  $originalsFolder"
Write-Host "  Copies (duplicates)  : $($stats.Copies)  →  $copiesFolder"
if ($Move)         { Write-Host "  Source files removed : $($stats.Deleted)  (Move mode)"   -ForegroundColor Yellow }
elseif ($DeleteCopies) { Write-Host "  Duplicates deleted   : $($stats.Deleted)  from source" -ForegroundColor Yellow }
if ($Compress)     { Write-Host "  Compressed           : $Compress  →  $OutputPath" }
if ($stats.Skipped -gt 0) { Write-Host "  Skipped (identical)  : $($stats.Skipped)"          -ForegroundColor Yellow }
if ($stats.Errors  -gt 0) { Write-Host "  Errors               : $($stats.Errors)"           -ForegroundColor Red }
Write-Host "========================================`n" -ForegroundColor Green
