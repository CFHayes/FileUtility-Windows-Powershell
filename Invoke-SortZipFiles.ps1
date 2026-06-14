<#
.SYNOPSIS
    Sorts compressed archives, installers, and disk images into a
    year\month folder structure based on the best available date.

.DESCRIPTION
    Date source priority (best → fallback):
      1. Windows Shell metadata  — embedded date properties via Shell COM API
      2. File LastWriteTime      — GUARANTEED fallback when metadata is absent
      3. Unresolved\             — only if the file system cannot provide any date

    Files are organised into three named subfolders within each year\month:
      Archives\    — compressed containers (ZIP, RAR, 7Z, TAR, GZ, etc.)
      Installers\  — setup packages and app bundles (MSI, MSIX, APK, JAR, etc.)
      DiskImages\  — mountable disk and VM images (ISO, IMG, VHD, VMDK, etc.)

    Output structure:
      OutputPath\
        2022\
          06 - June\
            Archives\
            Installers\
            DiskImages\
        2023\
          11 - November\
            Archives\
        Unresolved\   ← files where no date could be determined

    File types included:

      Archives   : .zip .rar .7z .tar .gz .bz2 .xz .tgz .tbz2 .lzma
                   .cab .z .lz .zst .br .lzh .ace .arj .sit .sitx
      Installers : .msi .msix .msixbundle .appx .appxbundle
                   .deb .rpm .pkg .dmg .apk .ipa .jar .war .ear .nupkg .vsix
      Disk Images: .iso .img .vhd .vhdx .vmdk .vdi .qcow2 .ova .ovf
                   .bin .mdf .mds .nrg .cue .toast

    Excluded by design:
      - EXE files (too broad — not all EXEs are installers)
      - All photo, video, and document extensions handled by other suite scripts
      - Windows system files (DLL, SYS, etc.)

    Requires Shared-Functions.ps1 in the same folder.

.PARAMETER SourcePath
    Folder containing the files to sort.

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
    Dry run — shows what would happen without touching any files.

.EXAMPLE
    # Copy, named month folders (default)
    .\Invoke-SortZipFiles.ps1 -SourcePath "C:\Downloads" -OutputPath "C:\Sorted"

.EXAMPLE
    # Move, recurse subfolders, numeric month folders
    .\Invoke-SortZipFiles.ps1 -SourcePath "C:\Downloads" -OutputPath "C:\Sorted" -Recurse -Move -FolderFormat Number

.EXAMPLE
    # Always dry-run first
    .\Invoke-SortZipFiles.ps1 -SourcePath "C:\Downloads" -OutputPath "C:\Sorted" -Recurse -WhatIf
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
# Load shared functions
# ---------------------------------------------------------------------------

. "$PSScriptRoot\Shared-Functions.ps1"

# ---------------------------------------------------------------------------
# Extension definitions
# Each maps to a subfolder name within year\month\
# ---------------------------------------------------------------------------

$ArchiveExtensions = @(
    '.zip', '.rar', '.7z', '.tar', '.gz',
    '.bz2', '.xz', '.tgz', '.tbz2', '.lzma',
    '.cab', '.z', '.lz', '.zst', '.br',
    '.lzh', '.ace', '.arj', '.sit', '.sitx'
)

$InstallerExtensions = @(
    '.msi', '.msix', '.msixbundle',
    '.appx', '.appxbundle',
    '.deb', '.rpm', '.pkg', '.dmg',
    '.apk', '.ipa',
    '.jar', '.war', '.ear',
    '.nupkg', '.vsix'
)

$DiskImageExtensions = @(
    '.iso', '.img', '.vhd', '.vhdx',
    '.vmdk', '.vdi', '.qcow2',
    '.ova', '.ovf',
    '.bin', '.mdf', '.mds', '.nrg',
    '.cue', '.toast'
)

# Combined list for file collection
$AllSupportedExtensions = $ArchiveExtensions + $InstallerExtensions + $DiskImageExtensions

# Subfolder name lookup by extension
$ExtensionCategoryMap = @{}
foreach ($ext in $ArchiveExtensions)    { $ExtensionCategoryMap[$ext] = 'Archives' }
foreach ($ext in $InstallerExtensions)  { $ExtensionCategoryMap[$ext] = 'Installers' }
foreach ($ext in $DiskImageExtensions)  { $ExtensionCategoryMap[$ext] = 'DiskImages' }

# ---------------------------------------------------------------------------
# Windows Shell COM API (same approach as Invoke-SortVideos.ps1)
# ---------------------------------------------------------------------------

$ShellApp = New-Object -ComObject Shell.Application

function Get-ShellDate {
    <#
    Queries Windows Shell property columns for an embedded file date.
    Falls through silently when the property is unavailable for a given
    file type — the caller handles the LastWriteTime fallback.

    Column indices used:
      208 = System.Media.DateEncoded   — most accurate embedded date
      191 = System.Document.DateCreated
      197 = System.Media.DateReleased
        4 = System.DateModified        — Shell-reported modified date
    #>
    param ([string]$FilePath)

    try {
        $folder   = $ShellApp.NameSpace([System.IO.Path]::GetDirectoryName($FilePath))
        $fileItem = $folder.ParseName([System.IO.Path]::GetFileName($FilePath))
        if ($null -eq $fileItem) { return $null }

        foreach ($idx in @(208, 191, 197)) {
            try {
                $val = $folder.GetDetailsOf($fileItem, $idx)
                if ($val -and $val.Trim() -ne '') {
                    # Strip hidden Unicode chars Shell sometimes inserts
                    $cleaned = ($val -replace '[^\x20-\x7E]', '').Trim()
                    $parsed  = $null
                    if ($cleaned -and [datetime]::TryParse(
                            $cleaned,
                            [System.Globalization.CultureInfo]::CurrentCulture,
                            [System.Globalization.DateTimeStyles]::None,
                            [ref]$parsed)) {
                        if ($parsed.Year -gt 1970 -and
                            $parsed.Year -le ([datetime]::Now.Year + 1)) {
                            $label = switch ($idx) {
                                208 { 'Shell: Media created' }
                                191 { 'Shell: Document date created' }
                                197 { 'Shell: Date released' }
                            }
                            return [PSCustomObject]@{ Date = $parsed; Source = $label }
                        }
                    }
                }
            }
            catch { <# Column not supported for this file type — try next #> }
        }
    }
    catch { <# Shell could not open file — fall through to LastWriteTime #> }

    return $null
}

function Get-BestDate {
    <#
    Returns the best available date for a file.

    Priority:
      1. Windows Shell embedded metadata  — any property column that holds a date
      2. File LastWriteTime               — guaranteed fallback via Shared-Functions.ps1
         (CreationTime is NOT used — Windows resets it on every file copy)
    #>
    param ([System.IO.FileInfo]$File)

    # Priority 1 — Shell metadata
    $shellResult = Get-ShellDate -FilePath $File.FullName
    if ($null -ne $shellResult) { return $shellResult }

    # Priority 2 — Guaranteed LastWriteTime fallback (from Shared-Functions.ps1)
    return Get-LastWriteTimeFallback -File $File
}

function Get-FileCategory {
    <#
    Returns the category subfolder name for a given file extension.
    Archives → 'Archives' | Installers → 'Installers' | Images → 'DiskImages'
    #>
    param ([string]$Extension)
    $ext = $Extension.ToLower()
    if ($ExtensionCategoryMap.ContainsKey($ext)) {
        return $ExtensionCategoryMap[$ext]
    }
    return 'Unresolved'
}

# ---------------------------------------------------------------------------
# Setup
# ---------------------------------------------------------------------------

Assert-SourcePath -Path $SourcePath

$unresolvedFolder = Join-Path $OutputPath 'Unresolved'
$modeLabel = if ($Move) { 'MOVE  (source files will be removed)' } else { 'COPY  (source untouched)' }

Write-Host "`nMode          : $modeLabel"    -ForegroundColor Yellow
Write-Host "Folder format : $FolderFormat"  -ForegroundColor Yellow
Write-Host "Output root   : $OutputPath"    -ForegroundColor Yellow
Write-Host "Metadata via  : Windows Shell API + LastWriteTime fallback`n" -ForegroundColor Yellow

# ---------------------------------------------------------------------------
# Step 1 — Collect files
# ---------------------------------------------------------------------------

Write-Host "[1/3] Scanning '$SourcePath'..." -ForegroundColor Cyan

$allFiles = Get-FilteredFiles `
    -FolderPath        $SourcePath `
    -Recurse:$Recurse `
    -IncludeExtensions $AllSupportedExtensions `
    -EmptyMessage      "  No supported archive, installer, or disk image files found in '$SourcePath'."

if ($allFiles.Count -eq 0) { exit 0 }

# Count by category for the scan summary
$archiveCount  = @($allFiles | Where-Object { $_.Extension.ToLower() -in $ArchiveExtensions   }).Count
$installCount  = @($allFiles | Where-Object { $_.Extension.ToLower() -in $InstallerExtensions }).Count
$imageCount    = @($allFiles | Where-Object { $_.Extension.ToLower() -in $DiskImageExtensions }).Count

Write-Host "      Found $($allFiles.Count) file(s):"
Write-Host "        Archives    : $archiveCount"
Write-Host "        Installers  : $installCount"
Write-Host "        Disk Images : $imageCount"

# ---------------------------------------------------------------------------
# Step 2 — Resolve best date per file
# ---------------------------------------------------------------------------

Write-Host "`n[2/3] Resolving dates..." -ForegroundColor Cyan

$resolved   = [System.Collections.Generic.List[PSCustomObject]]::new()
$unresolved = [System.Collections.Generic.List[System.IO.FileInfo]]::new()

$total = $allFiles.Count
$index = 0

foreach ($file in $allFiles) {
    $index++
    Write-Progress -Activity 'Reading metadata' `
                   -Status   "$index / $total : $($file.Name)" `
                   -PercentComplete (($index / $total) * 100)

    try {
        $bestDate = Get-BestDate -File $file

        if ($null -ne $bestDate) {
            $resolved.Add([PSCustomObject]@{
                File      = $file
                Date      = $bestDate.Date
                DateSource= $bestDate.Source
                Category  = Get-FileCategory -Extension $file.Extension
            })
        }
        else {
            $unresolved.Add($file)
        }
    }
    catch {
        Write-Warning "Could not resolve date for '$($file.FullName)': $_"
        $unresolved.Add($file)
    }
}
Write-Progress -Activity 'Reading metadata' -Completed

$shellCount    = @($resolved | Where-Object { $_.DateSource -like 'Shell:*'    }).Count
$fallbackCount = @($resolved | Where-Object { $_.DateSource -like '*fallback*' }).Count

Write-Host "      Shell metadata found : $shellCount file(s)"
Write-Host "      Fallback date used   : $fallbackCount file(s)"
Write-Host "      No date (Unresolved) : $($unresolved.Count) file(s)"

# ---------------------------------------------------------------------------
# Step 3 — Sort into year\month\category folders
# ---------------------------------------------------------------------------

Write-Host "`n[3/3] Sorting files..." -ForegroundColor Cyan

$stats = @{
    Copied     = 0
    Moved      = 0
    Skipped    = 0
    Unresolved = 0
    Errors     = 0
}

$index = 0
$total = $resolved.Count

foreach ($item in $resolved) {
    $index++
    Write-Progress -Activity 'Sorting' `
                   -Status   "$index / $total : $($item.File.Name)" `
                   -PercentComplete (($index / $total) * 100)

    try {
        # Structure: OutputPath\yyyy\mm\Category\
        $yearFolder  = Join-Path $OutputPath $item.Date.Year.ToString()
        $monthFolder = Join-Path $yearFolder (Get-MonthFolderName -Date $item.Date -Format $FolderFormat)
        $destFolder  = Join-Path $monthFolder $item.Category

        $result = Invoke-FileTransfer `
            -FilePath     $item.File.FullName `
            -DestFolder   $destFolder `
            -DeleteSource $Move.IsPresent

        switch -Wildcard ($result.Action) {
            'Moved'    { $stats.Moved++ }
            'Copied'   { $stats.Copied++ }
            'Skipped*' { $stats.Skipped++ }
        }

        Write-Verbose "$($result.Action): '$($item.File.Name)' [$($item.Category) | $($item.DateSource)] → $($result.Path)"
    }
    catch {
        Write-Warning "Failed to process '$($item.File.FullName)': $_"
        $stats.Errors++
    }
}

# Unresolved files → Unresolved\ folder
foreach ($file in $unresolved) {
    try {
        Invoke-FileTransfer `
            -FilePath     $file.FullName `
            -DestFolder   $unresolvedFolder `
            -DeleteSource $Move.IsPresent | Out-Null
        $stats.Unresolved++
    }
    catch {
        Write-Warning "Failed to copy unresolved '$($file.FullName)': $_"
        $stats.Errors++
    }
} 

Write-Progress -Activity 'Sorting' -Completed

# Cleanup COM object
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ShellApp) | Out-Null

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "  Archive / Installer Sort Complete"      -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Mode              : $modeLabel"
Write-Host "  Files found       : $($allFiles.Count)"
Write-Host "    Archives        : $archiveCount"
Write-Host "    Installers      : $installCount"
Write-Host "    Disk Images     : $imageCount"
Write-Host "  Shell metadata    : $shellCount"
Write-Host "  Fallback dates    : $fallbackCount"
if ($Move) { Write-Host "  Moved             : $($stats.Moved)" }
else       { Write-Host "  Copied            : $($stats.Copied)" }
if ($stats.Skipped    -gt 0) { Write-Host "  Skipped (identical): $($stats.Skipped)"              -ForegroundColor Yellow }
if ($stats.Unresolved -gt 0) { Write-Host "  Unresolved         : $($stats.Unresolved)  →  $unresolvedFolder" -ForegroundColor Yellow }
if ($stats.Errors     -gt 0) { Write-Host "  Errors             : $($stats.Errors)"               -ForegroundColor Red }
Write-Host "========================================`n" -ForegroundColor Green
