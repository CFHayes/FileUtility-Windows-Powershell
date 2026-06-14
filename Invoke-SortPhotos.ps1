<#
.SYNOPSIS
    Sorts image files into year\month folder structure based on the most
    accurate available date, preferring EXIF DateTimeOriginal.

.DESCRIPTION
    Date source priority (best → fallback):
      1. EXIF DateTimeOriginal  — set by the camera at the moment of capture
      2. EXIF DateTimeDigitized — set when image was digitised (same as #1 on most cameras)
      3. File LastWriteTime     — last modified date (used only if EXIF absent)
      4. Unresolved\            — folder for files where no reliable date can be found

    Output structure:
      OutputPath\
        2022\
          01 - January\
          06 - June\
        2023\
          11 - November\
        Unresolved\   ← files with no EXIF and no reliable fallback date

    Supported formats: JPG, JPEG, PNG, TIFF, TIF, HEIC, HEIF, BMP, GIF, WEBP, CR2,
                       CR3, NEF, ARW, ORF, RW2, DNG, RAF (RAW camera formats)

.PARAMETER SourcePath
    Folder containing the image files to sort.

.PARAMETER OutputPath
    Root folder where the year\month structure will be created.

.PARAMETER Recurse
    Scan subfolders of SourcePath as well.

.PARAMETER Move
    Move files instead of copying them. Source file is removed after a
    successful transfer.

.PARAMETER FolderFormat
    Controls the month subfolder naming style.
      Named    →  06 - June          (default)
      Number   →  06
      Full     →  2024-06

.PARAMETER WhatIf
    Dry run — shows what would happen without moving or copying anything.

.EXAMPLE
    # Copy photos, named month folders
    .\Invoke-SortPhotos.ps1 -SourcePath "C:\Photos" -OutputPath "C:\Sorted"

.EXAMPLE
    # Move photos, recurse subfolders, numeric month folders
    .\Invoke-SortPhotos.ps1 -SourcePath "C:\Photos" -OutputPath "C:\Sorted" -Recurse -Move -FolderFormat Number

.EXAMPLE
    # Dry run first to check results before committing
    .\Invoke-SortPhotos.ps1 -SourcePath "C:\Photos" -OutputPath "C:\Sorted" -Recurse -Move -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory)][string] $SourcePath,
    [Parameter(Mandatory)][string] $OutputPath,
    [switch] $Recurse,
    [switch] $Move,
    [ValidateSet('Named', 'Number', 'Full')]
    [string] $FolderFormat = 'Named'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

$SupportedExtensions = @(
    '.jpg', '.jpeg', '.png', '.tiff', '.tif',
    '.heic', '.heif', '.bmp', '.gif', '.webp',
    '.cr2', '.cr3', '.nef', '.arw', '.orf',
    '.rw2', '.dng', '.raf'
)

$MonthNames = @{
    1='January'; 2='February'; 3='March';    4='April';
    5='May';     6='June';     7='July';      8='August';
    9='September'; 10='October'; 11='November'; 12='December'
}

# EXIF tag IDs we care about (as decimal)
# 36867 = DateTimeOriginal
# 36868 = DateTimeDigitized
$ExifDateTags = @(36867, 36868)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Get-ExifDate {
    <#
    Reads EXIF date tags from an image file using System.Drawing.
    Returns a [datetime] or $null if no EXIF date is found.
    EXIF date format: "yyyy:MM:dd HH:mm:ss"
    #>
    param ([string]$FilePath)

    try {
        Add-Type -AssemblyName System.Drawing

        $img = [System.Drawing.Image]::FromFile($FilePath)
        try {
            $propIds = $img.PropertyIdList

            foreach ($tagId in $ExifDateTags) {
                if ($tagId -notin $propIds) { continue }

                $prop  = $img.GetPropertyItem($tagId)
                # Property is ASCII bytes; strip null terminator
                $raw   = [System.Text.Encoding]::ASCII.GetString($prop.Value).TrimEnd([char]0)

                # Validate it looks like a real date (some cameras write all zeros)
                if ($raw -match '^\d{4}:\d{2}:\d{2} \d{2}:\d{2}:\d{2}$' -and
                    $raw -notmatch '^0000') {
                    $dt = [datetime]::ParseExact($raw, 'yyyy:MM:dd HH:mm:ss',
                                                 [System.Globalization.CultureInfo]::InvariantCulture)
                    return $dt
                }
            }
        }
        finally { $img.Dispose() }
    }
    catch {
        # System.Drawing can't open some RAW formats — that's expected; fall through
    }

    return $null
}

function Get-BestDate {
    <#
    Returns the best available date for a file and a label describing the source.

    Priority:
      1. EXIF DateTimeOriginal  — camera capture time (most accurate)
      2. EXIF DateTimeDigitized — digitisation time (same as #1 on most cameras)
      3. File LastWriteTime     — GUARANTEED fallback when EXIF is absent or stripped.
                                  A file always has a LastWriteTime so this never
                                  returns $null unless the file itself is unreadable.

    Note: LastWriteTime is preferred over CreationTime as a fallback because
    CreationTime is reset by Windows every time a file is copied, whereas
    LastWriteTime is preserved across most copy operations.
    #>
    param ([System.IO.FileInfo]$File)

    # ── Priority 1 & 2: EXIF ──────────────────────────────────────────────
    $exifDate = Get-ExifDate -FilePath $File.FullName
    if ($null -ne $exifDate) {
        return [PSCustomObject]@{ Date = $exifDate; Source = 'EXIF DateTimeOriginal' }
    }

    # ── Priority 3: File LastWriteTime (guaranteed fallback) ──────────────
    # This branch is always reached when EXIF is absent, stripped, or unreadable.
    # We do NOT gate this on a year range — any valid LastWriteTime is accepted
    # so that no file is ever sent to Unresolved\ simply due to an unusual date.
    $lwt = $File.LastWriteTime
    if ($null -ne $lwt -and $lwt -ne [datetime]::MinValue) {
        return [PSCustomObject]@{ Date = $lwt; Source = 'LastWriteTime (EXIF absent — fallback)' }
    }

    # ── Unresolvable — file system couldn't return any date at all ─────────
    # In practice this should never be reached on a readable Windows file.
    return $null
}

function Get-MonthFolderName {
    param ([datetime]$Date)

    switch ($FolderFormat) {
        'Number' { return '{0:D2}' -f $Date.Month }
        'Full'   { return '{0}-{1:D2}' -f $Date.Year, $Date.Month }
        default  { return '{0:D2} - {1}' -f $Date.Month, $MonthNames[$Date.Month] }
    }
}

function Copy-ToDestination {
    param (
        [string] $FilePath,
        [string] $DestFolder,
        [bool]   $DeleteSource
    )

    if (-not (Test-Path $DestFolder)) {
        if ($PSCmdlet.ShouldProcess($DestFolder, 'Create directory')) {
            New-Item -ItemType Directory -Path $DestFolder -Force | Out-Null
        }
    }

    $fileName  = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    $extension = [System.IO.Path]::GetExtension($FilePath)
    $destPath  = Join-Path $DestFolder "$fileName$extension"

    # Collision handling — skip if file is identical (same size + name), rename otherwise
    $counter = 1
    while (Test-Path $destPath) {
        $existing = Get-Item $destPath
        $incoming = Get-Item $FilePath
        if ($existing.Length -eq $incoming.Length) {
            # Very likely the same file already copied — skip
            return [PSCustomObject]@{ Path = $destPath; Action = 'Skipped (already exists)' }
        }
        $destPath = Join-Path $DestFolder "${fileName}_$counter$extension"
        $counter++
    }

    if ($PSCmdlet.ShouldProcess($destPath, "$(if ($DeleteSource) {'Move'} else {'Copy'}) '$FilePath'")) {
        Copy-Item -LiteralPath $FilePath -Destination $destPath -Force
        if ($DeleteSource) {
            Remove-Item -LiteralPath $FilePath -Force
        }
    }

    return [PSCustomObject]@{ Path = $destPath; Action = if ($DeleteSource) { 'Moved' } else { 'Copied' } }
}

# ---------------------------------------------------------------------------
# Setup
# ---------------------------------------------------------------------------

$unresolvedFolder = Join-Path $OutputPath 'Unresolved'

$modeLabel = if ($Move) { 'MOVE  (source files will be removed)' } else { 'COPY  (source untouched)' }
Write-Host "`nMode          : $modeLabel"          -ForegroundColor Yellow
Write-Host "Folder format : $FolderFormat"         -ForegroundColor Yellow
Write-Host "Output root   : $OutputPath`n"         -ForegroundColor Yellow

# ---------------------------------------------------------------------------
# Step 1 — Collect image files
# ---------------------------------------------------------------------------

Write-Host "[1/3] Scanning '$SourcePath'..." -ForegroundColor Cyan

$getChildParams = @{ LiteralPath = $SourcePath; File = $true }
if ($Recurse) { $getChildParams['Recurse'] = $true }

# @() wraps the result in an array — prevents .Count failing when the
# folder is empty or contains no matching files (PowerShell returns $null
# from a pipeline with no results, and $null has no .Count property)
$allFiles = @(Get-ChildItem @getChildParams |
    Where-Object { $_.Extension.ToLower() -in $SupportedExtensions })

Write-Host "      Found $($allFiles.Count) supported image file(s)."

if ($allFiles.Count -eq 0) {
    Write-Host "`n  No supported image files found in '$SourcePath'." -ForegroundColor Yellow
    Write-Host "  Supported extensions: $($SupportedExtensions -join ', ')`n" -ForegroundColor Yellow
    exit 0
}

# ---------------------------------------------------------------------------
# Step 2 — Resolve best date for each file
# ---------------------------------------------------------------------------

Write-Host "`n[2/3] Resolving dates..." -ForegroundColor Cyan

$resolved   = [System.Collections.Generic.List[PSCustomObject]]::new()
$unresolved = [System.Collections.Generic.List[System.IO.FileInfo]]::new()

$total = $allFiles.Count
$index = 0

foreach ($file in $allFiles) {
    $index++
    Write-Progress -Activity 'Reading dates' `
                   -Status   "$index / $total : $($file.Name)" `
                   -PercentComplete (($index / $total) * 100)

    $bestDate = Get-BestDate -File $file

    if ($null -ne $bestDate) {
        $resolved.Add([PSCustomObject]@{
            File       = $file
            Date       = $bestDate.Date
            DateSource = $bestDate.Source
        })
    }
    else {
        $unresolved.Add($file)
    }
}

Write-Progress -Activity 'Reading dates' -Completed

$exifCount      = @($resolved | Where-Object { $_.DateSource -eq 'EXIF DateTimeOriginal' }).Count
$fallbackCount  = @($resolved | Where-Object { $_.DateSource -like '*fallback*' }).Count

Write-Host "      EXIF date found       : $exifCount file(s)"
Write-Host "      Fallback date used    : $fallbackCount file(s)"
Write-Host "      No date (Unresolved)  : $($unresolved.Count) file(s)"

# ---------------------------------------------------------------------------
# Step 3 — Copy / Move files into year\month structure
# ---------------------------------------------------------------------------

Write-Host "`n[3/3] Sorting files..." -ForegroundColor Cyan

$stats = @{
    Copied      = 0
    Moved       = 0
    Skipped     = 0
    Unresolved  = 0
    Errors      = 0
}

# Resolved files → year\month folders
$index = 0
$total = $resolved.Count

foreach ($item in $resolved) {
    $index++
    Write-Progress -Activity 'Sorting' `
                   -Status   "$index / $total : $($item.File.Name)" `
                   -PercentComplete (($index / $total) * 100)

    try {
        $yearFolder  = Join-Path $OutputPath $item.Date.Year.ToString()
        $monthFolder = Join-Path $yearFolder  (Get-MonthFolderName -Date $item.Date)

        $result = Copy-ToDestination -FilePath     $item.File.FullName `
                                     -DestFolder   $monthFolder `
                                     -DeleteSource $Move.IsPresent

        switch ($result.Action) {
            'Moved'                    { $stats.Moved++ }
            'Copied'                   { $stats.Copied++ }
            'Skipped (already exists)' { $stats.Skipped++ }
        }

        # Verbose log per file
        Write-Verbose "$($result.Action): '$($item.File.Name)' [$($item.DateSource)] → $($result.Path)"
    }
    catch {
        Write-Warning "Failed to process '$($item.File.FullName)': $_"
        $stats.Errors++
    }
}

# Unresolved files → Unresolved\ folder
foreach ($file in $unresolved) {
    try {
        Copy-ToDestination -FilePath     $file.FullName `
                           -DestFolder   $unresolvedFolder `
                           -DeleteSource $Move.IsPresent | Out-Null
        $stats.Unresolved++
    }
    catch {
        Write-Warning "Failed to copy unresolved file '$($file.FullName)': $_"
        $stats.Errors++
    }
}

Write-Progress -Activity 'Sorting' -Completed

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "  Photo Sort Complete" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Mode              : $modeLabel"
Write-Host "  Images found      : $($allFiles.Count)"
Write-Host "  EXIF dates        : $exifCount"
Write-Host "  Fallback dates    : $fallbackCount"
if ($Move) {
    Write-Host "  Moved             : $($stats.Moved)"
} else {
    Write-Host "  Copied            : $($stats.Copied)"
}
if ($stats.Skipped -gt 0) {
    Write-Host "  Skipped (dupes)   : $($stats.Skipped)"   -ForegroundColor Yellow
}
if ($stats.Unresolved -gt 0) {
    Write-Host "  Unresolved        : $($stats.Unresolved)  →  $unresolvedFolder" -ForegroundColor Yellow
}
if ($stats.Errors -gt 0) {
    Write-Host "  Errors            : $($stats.Errors)"    -ForegroundColor Red
}
Write-Host "========================================`n" -ForegroundColor Green
