<#
.SYNOPSIS
    Sorts video files into year\month folder structure based on the most
    accurate available date, using the Windows Shell COM API.
    No third-party tools required.

.DESCRIPTION
    Date source priority (best → fallback):
      1. Shell: Media created (DateEncoded)  — recording date embedded by camera/device
      2. Shell: Date encoded                 — encoding timestamp
      3. Shell: Date released                — release/publish date tag
      4. File LastWriteTime                  — guaranteed fallback when metadata absent
      5. Unresolved\                         — only if file system cannot provide any date

    All metadata is read via the Windows Shell COM API (the same engine that
    powers File Explorer's Details pane). No third-party tools required.

    Requires Shared-Functions.ps1 in the same folder.

.PARAMETER SourcePath
    Folder containing the video files to sort.

.PARAMETER OutputPath
    Root folder where the year\month structure will be created.

.PARAMETER Recurse
    Scan subfolders of SourcePath as well.

.PARAMETER Move
    Move files instead of copying them.

.PARAMETER FolderFormat
    Named  →  06 - June  (default)
    Number →  06
    Full   →  2024-06

.PARAMETER WhatIf
    Dry run — shows what would happen without touching any files.

.EXAMPLE
    .\Invoke-SortVideos.ps1 -SourcePath "C:\Videos" -OutputPath "C:\Sorted"

.EXAMPLE
    .\Invoke-SortVideos.ps1 -SourcePath "C:\Videos" -OutputPath "C:\Sorted" -Recurse -Move -WhatIf
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
# Constants
# ---------------------------------------------------------------------------

$SupportedExtensions = @(
    '.mp4', '.mov', '.m4v', '.avi', '.mkv',
    '.wmv', '.flv', '.webm', '.mpg', '.mpeg',
    '.3gp', '.3g2', '.mts', '.m2ts', '.ts',
    '.vob', '.asf', '.dv'
)

# ---------------------------------------------------------------------------
# Video-specific date helper (Windows Shell COM API)
# ---------------------------------------------------------------------------

$ShellApp = New-Object -ComObject Shell.Application

function Get-ShellDate {
    <#
    Queries Windows Shell property columns for video recording date.
    Column indices:
      208 = Media created (System.Media.DateEncoded) — most accurate for video
      191 = Date encoded
      197 = Date released
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
                    $cleaned = $val -replace '[^\x20-\x7E]', '' | ForEach-Object { $_.Trim() }
                    $parsed  = $null
                    if ($cleaned -and [datetime]::TryParse($cleaned,
                            [System.Globalization.CultureInfo]::CurrentCulture,
                            [System.Globalization.DateTimeStyles]::None,
                            [ref]$parsed)) {
                        if ($parsed.Year -gt 1970 -and $parsed.Year -le ([datetime]::Now.Year + 1)) {
                            $sourceLabel = switch ($idx) {
                                208 { 'Shell: Media created (DateEncoded)' }
                                191 { 'Shell: Date encoded' }
                                197 { 'Shell: Date released' }
                            }
                            return [PSCustomObject]@{ Date = $parsed; Source = $sourceLabel }
                        }
                    }
                }
            }
            catch { <# Column not available for this file type — try next #> }
        }
    }
    catch { <# Shell couldn't open file — fall through #> }

    return $null
}

function Get-BestDate {
    param ([System.IO.FileInfo]$File)

    $shellResult = Get-ShellDate -FilePath $File.FullName
    if ($null -ne $shellResult) { return $shellResult }

    # Guaranteed fallback from Shared-Functions.ps1
    return Get-LastWriteTimeFallback -File $File
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
Write-Host "Metadata via  : Windows Shell API (no third-party tools needed)`n" -ForegroundColor Yellow

# ---------------------------------------------------------------------------
# Step 1 — Collect video files
# ---------------------------------------------------------------------------

Write-Host "[1/3] Scanning '$SourcePath'..." -ForegroundColor Cyan

$allFiles = Get-FilteredFiles -FolderPath         $SourcePath `
                              -Recurse:$Recurse `
                              -IncludeExtensions   $SupportedExtensions `
                              -EmptyMessage        "  No supported video files found in '$SourcePath'."

if ($allFiles.Count -eq 0) { exit 0 }
Write-Host "      Found $($allFiles.Count) supported video file(s)."

# ---------------------------------------------------------------------------
# Step 2 — Resolve best date per file
# ---------------------------------------------------------------------------

Write-Host "`n[2/3] Resolving dates (Windows Shell metadata)..." -ForegroundColor Cyan

$resolved   = [System.Collections.Generic.List[PSCustomObject]]::new()
$unresolved = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
$total = $allFiles.Count
$index = 0

foreach ($file in $allFiles) {
    $index++
    Write-Progress -Activity 'Reading metadata' `
                   -Status   "$index / $total : $($file.Name)" `
                   -PercentComplete (($index / $total) * 100)

    $bestDate = Get-BestDate -File $file
    if ($null -ne $bestDate) {
        $resolved.Add([PSCustomObject]@{ File = $file; Date = $bestDate.Date; DateSource = $bestDate.Source })
    }
    else { $unresolved.Add($file) }
}
Write-Progress -Activity 'Reading metadata' -Completed

$shellCount    = @($resolved | Where-Object { $_.DateSource -like 'Shell:*' }).Count
$fallbackCount = @($resolved | Where-Object { $_.DateSource -like '*fallback*' }).Count

Write-Host "      Shell metadata found   : $shellCount file(s)"
Write-Host "      Fallback date used     : $fallbackCount file(s)"
Write-Host "      No date (Unresolved)   : $($unresolved.Count) file(s)"

# ---------------------------------------------------------------------------
# Step 3 — Sort into year\month folders
# ---------------------------------------------------------------------------

Write-Host "`n[3/3] Sorting files..." -ForegroundColor Cyan

$stats = @{ Copied = 0; Moved = 0; Skipped = 0; Unresolved = 0; Errors = 0 }
$index = 0
$total = $resolved.Count

foreach ($item in $resolved) {
    $index++
    Write-Progress -Activity 'Sorting' `
                   -Status   "$index / $total : $($item.File.Name)" `
                   -PercentComplete (($index / $total) * 100)
    try {
        $destFolder = Join-Path (Join-Path $OutputPath $item.Date.Year.ToString()) `
                                (Get-MonthFolderName -Date $item.Date -Format $FolderFormat)

        $result = Invoke-FileTransfer -FilePath     $item.File.FullName `
                                      -DestFolder   $destFolder `
                                      -DeleteSource $Move.IsPresent

        switch -Wildcard ($result.Action) {
            'Moved'    { $stats.Moved++ }
            'Copied'   { $stats.Copied++ }
            'Skipped*' { $stats.Skipped++ }
        }
        Write-Verbose "$($result.Action): '$($item.File.Name)' [$($item.DateSource)] → $($result.Path)"
    }
    catch {
        Write-Warning "Failed to process '$($item.File.FullName)': $_"
        $stats.Errors++
    }
}

foreach ($file in $unresolved) {
    try {
        Invoke-FileTransfer -FilePath $file.FullName -DestFolder $unresolvedFolder `
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
Write-Host "  Video Sort Complete"                     -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Mode              : $modeLabel"
Write-Host "  Videos found      : $($allFiles.Count)"
Write-Host "  Shell metadata    : $shellCount"
Write-Host "  Fallback dates    : $fallbackCount"
if ($Move) { Write-Host "  Moved   : $($stats.Moved)" }
else       { Write-Host "  Copied  : $($stats.Copied)" }
if ($stats.Skipped    -gt 0) { Write-Host "  Skipped (identical)  : $($stats.Skipped)"              -ForegroundColor Yellow }
if ($stats.Unresolved -gt 0) { Write-Host "  Unresolved           : $($stats.Unresolved)  →  $unresolvedFolder" -ForegroundColor Yellow }
if ($stats.Errors     -gt 0) { Write-Host "  Errors               : $($stats.Errors)"               -ForegroundColor Red }
Write-Host "========================================`n" -ForegroundColor Green
