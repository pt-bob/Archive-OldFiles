# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

SP-Archive-XML is a PowerShell project for archiving SalesPad live "In" and "Out" XML files to another location.

## Development

This is a PowerShell-based project. Scripts should be developed and tested using PowerShell 5.1+ or PowerShell Core 7+.

## Scripts

### Archive-OldFiles.ps1

Archives files older than a specified year from a source directory to a destination directory. Uses robocopy for reliable file copying.

**Parameters:**
- `-DryRun` - Preview files to be copied without executing (no confirmation prompt)
- `-Recurse` - Include subdirectories (default: root folder only)
- `-Help` - Display usage information

**Usage:**
```powershell
# Interactive mode (prompts for source, destination, extension, year)
.\Archive-OldFiles.ps1

# Preview only
.\Archive-OldFiles.ps1 -DryRun

# Include subdirectories
.\Archive-OldFiles.ps1 -Recurse

# Show help
.\Archive-OldFiles.ps1 -Help
```

**Interactive Prompts:**
- Source directory (default: `C:\SalesPad\XML`)
- Destination directory (default: `C:\Archive\SalesPad\XML`)
- File extension (default: `*.xml`)
- Cutoff year (leave blank to copy all files)

**Features:**
- Handles file paths with spaces
- Shows preview with file count, total size, and first 10 files before execution
- Displays the robocopy command that will be executed
- Requires confirmation before copying
- Logs operations to `Archive-OldFiles.log` in the script directory
- Copies files (preserves originals in source)

**Robocopy Parameters Used:**
- `/MINAGE:date` - Only copy files older than the specified date
- `/E` - Include subdirectories (only when `-Recurse` is used)
- `/R:3` - Retry 3 times on failed copies
- `/W:5` - Wait 5 seconds between retries
- `/NP` - No progress percentage
- `/LOG+:file` - Append to log file
