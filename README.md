# FileUtility-Windows-Powershell

A suite of Windows PowerShell scripts for intelligent file management and organization. Includes deduplication, photo/video sorting by date, document organization, and ZIP file management.

## 📋 Scripts Overview

### **Invoke-Dedups.ps1**
Finds and separates duplicate files using a three-tier validation approach:
- **Stage 1:** Group by file size + extension (fast metadata filter)
- **Stage 2:** MD5 hash of first/last 64KB (cheap secondary filter)
- **Stage 3:** Full SHA256 hash (definitive duplicate detection)

**Output:** Organizes files into `originals/` and `copies/` folders.

**Options:**
- `-SourcePath` — Folder containing files to deduplicate
- `-OutputPath` — Where to place results
- `-Recurse` — Include subfolders
- `-Move` — Move files instead of copying (removes from source)
- `-DeleteCopies` — Delete duplicates from source after copying
- `-Compress` — Compress results: `Source`, `Copies`, or `Both`
- `-WhatIf` — Dry run mode

**Example:**
```powershell
.\Invoke-Dedups.ps1 -SourcePath "C:\MyFiles" -OutputPath "C:\Sorted" -Recurse -DeleteCopies
```

---

### **Invoke-SortPhotos.ps1**
Organizes photos into date-based folder structures using EXIF metadata.

**Output:** Creates folders like `2024\06 - June\` based on photo capture date.

**Features:**
- Extracts date from photo EXIF data
- Falls back to file LastWriteTime if EXIF is absent
- Handles corrupted metadata gracefully

---

### **Invoke-SortVideos.ps1**
Similar to photo sorting but for video files, organizing by date created/modified.

**Output:** Date-organized video folders matching photo structure.

---

### **Invoke-Documents.ps1**
Organizes documents (PDFs, Word, Excel, etc.) by creation/modification date.

**Output:** Monthly folder hierarchy for easy archival and retrieval.

---

### **Invoke-SortZipFiles.ps1**
Extracts and organizes contents from ZIP archives into date-based directories.

**Features:**
- Extracts ZIP contents intelligently
- Applies photo/video/document sorting to extracted files
- Optional cleanup of original ZIP after extraction

---

### **Remove-0ByteFiles.ps1**
Utility script to remove empty (0-byte) files from a directory.

**Usage:**
```powershell
.\Remove-0ByteFiles.ps1 -SourcePath "C:\MyFiles" -Recurse
```

---

### **Shared-Functions.ps1**
**Library file — do not run directly.** Provides shared utilities used by all scripts:

- **File Transfer:** `Invoke-FileTransfer` — copy/move with SHA256 collision detection
- **Compression:** `Compress-OutputFolder` — zip folders with timestamps
- **Hashing:** `Get-PartialHash`, `Get-FullHash` — efficient duplicate detection
- **Date Utilities:** `Get-LastWriteTimeFallback`, `Get-MonthFolderName`
- **File Collection:** `Get-FilteredFiles` — scan folders with filtering
- **Validation:** `Assert-SourcePath` — validate input paths

---

## 🚀 Getting Started

### Requirements
- Windows 7 or later
- PowerShell 3.0 or later
- Administrator privileges recommended (for accessing restricted folders)

### Installation
1. Clone or download the repository
2. Keep all `.ps1` files in the same folder (scripts use relative paths to find `Shared-Functions.ps1`)
3. Run scripts from PowerShell with appropriate permissions

### Basic Usage

**Deduplicate a folder:**
```powershell
cd C:\path\to\scripts
.\Invoke-Dedups.ps1 -SourcePath "C:\MyPhotos" -OutputPath "C:\Sorted"
```

**Sort photos by date:**
```powershell
.\Invoke-SortPhotos.ps1 -SourcePath "C:\RawPhotos" -OutputPath "C:\OrganizedPhotos" -Recurse
```

**Dry run (preview changes):**
```powershell
.\Invoke-Dedups.ps1 -SourcePath "C:\MyFiles" -OutputPath "C:\Sorted" -WhatIf
```

---

## 🔧 Advanced Features

### Collision Resolution
When files with the same name already exist in the destination:
1. **Different size:** Renamed with `_1`, `_2` suffixes (files are different)
2. **Same size, same SHA256:** Skipped (confirmed duplicate)
3. **Same size, different SHA256:** Renamed (different content, coincidental size match)

This three-stage approach prevents data loss from false size-only comparisons.

### Compression
Create timestamped ZIP archives automatically:
```powershell
.\Invoke-Dedups.ps1 -SourcePath "C:\MyFiles" -OutputPath "C:\Sorted" -Compress Both
```
Output: `source_20240614T221342Z.zip` and `copies_20240614T221342Z.zip`

### Metadata Fallback
Photo/video sorters use intelligent date resolution:
1. Try embedded EXIF metadata
2. Fall back to file LastWriteTime (preserved across most copy operations)
3. Ignore invalid dates — no files discarded due to unusual timestamps

---

## 📊 Performance

- **Deduplication:** Uses three-tier hashing for efficient duplicate detection
  - Partial hashing (64KB) eliminates ~90% of false positives
  - Full SHA256 only on actual candidates
- **Large Datasets:** Supports files of any size with streaming hash calculation
- **Progress Reporting:** Real-time progress bars on long operations

---

## 🛡️ Safety

- **No data loss:** Copy-first approach with collision detection
- **WhatIf support:** All scripts support `-WhatIf` for dry runs
- **Source preservation:** By default, source files remain untouched
- **Error handling:** Verbose error reporting without stopping on individual failures

---

## 🐛 Troubleshooting

**"Could not partial-hash" warnings:**
- File may be in use or have permission issues
- Check file permissions and ensure it's not locked by another process

**No files found:**
- Verify source path exists: `Test-Path "C:\YourPath"`
- Check file extensions are supported
- Use `-Recurse` to search subfolders

**Duplicate detection not working:**
- Ensure both files have identical content (not just name/size)
- Check file permissions allow reading for hashing

---

## 📝 License

This project is provided as-is for Windows file management automation.

---

## 💡 Tips

- **Backup first:** Always test on a small subset before processing large collections
- **Use WhatIf:** Preview results with `-WhatIf` before committing to changes
- **Schedule tasks:** Use Windows Task Scheduler to automate regular organization
- **Combine operations:** Use `-Compress` to archive results for long-term storage

---

## 🤝 Contributing

Found a bug or have a suggestion? Feel free to open an issue or submit a pull request.
