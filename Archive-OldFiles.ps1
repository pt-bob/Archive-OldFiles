<#
.SYNOPSIS
    Archives files older than a specified year to a destination directory.

.DESCRIPTION
    Prompts for source/destination directories, file extension, and cutoff year.
    Shows a preview of the operation before executing.
    By default, only searches the root of the source folder.

.PARAMETER DryRun
    Shows what files would be copied without prompting for confirmation or executing.

.PARAMETER Recurse
    Search subdirectories in addition to the root folder.

.EXAMPLE
    .\Archive-OldFiles.ps1

.EXAMPLE
    .\Archive-OldFiles.ps1 -DryRun

.EXAMPLE
    .\Archive-OldFiles.ps1 -Recurse
#>

param(
    [switch]$DryRun,
    [switch]$Recurse,
    [switch]$Help
)

# Display help and exit
if ($Help) {
    Write-Host ""
    Write-Host "Archive-OldFiles.ps1" -ForegroundColor Cyan
    Write-Host "Archives files older than a specified year to a destination directory." -ForegroundColor Gray
    Write-Host ""
    Write-Host "USAGE:" -ForegroundColor Yellow
    Write-Host "  .\Archive-OldFiles.ps1 [-DryRun] [-Recurse] [-Help]"
    Write-Host ""
    Write-Host "PARAMETERS:" -ForegroundColor Yellow
    Write-Host "  -DryRun    Show preview of files to be copied without executing."
    Write-Host "             No confirmation prompt, no files copied."
    Write-Host ""
    Write-Host "  -Recurse   Include subdirectories when searching for files."
    Write-Host "             By default, only the root folder is searched."
    Write-Host ""
    Write-Host "  -Help      Display this help message."
    Write-Host ""
    Write-Host "EXAMPLES:" -ForegroundColor Yellow
    Write-Host "  .\Archive-OldFiles.ps1" -ForegroundColor White
    Write-Host "      Run interactively, search root folder only."
    Write-Host ""
    Write-Host "  .\Archive-OldFiles.ps1 -DryRun" -ForegroundColor White
    Write-Host "      Preview what would be copied without executing."
    Write-Host ""
    Write-Host "  .\Archive-OldFiles.ps1 -Recurse" -ForegroundColor White
    Write-Host "      Include subdirectories in the search."
    Write-Host ""
    Write-Host "  .\Archive-OldFiles.ps1 -DryRun -Recurse" -ForegroundColor White
    Write-Host "      Preview with subdirectories included."
    Write-Host ""
    exit 0
}

# Prompt for source directory
$defaultSource = "C:\SalesPad\XML"
$sourceDir = Read-Host "Enter source directory [$defaultSource]"
if ([string]::IsNullOrWhiteSpace($sourceDir)) {
    $sourceDir = $defaultSource
}

# Clean up path: trim whitespace and trailing backslash, remove surrounding quotes if user added them
$sourceDir = $sourceDir.Trim().Trim('"').Trim("'").TrimEnd('\')

# Validate source directory
if (-not (Test-Path -LiteralPath $sourceDir -PathType Container)) {
    Write-Host "ERROR: Source directory does not exist: $sourceDir" -ForegroundColor Red
    Write-Host "       (If you copied this path, verify it exists in Explorer)" -ForegroundColor DarkGray
    exit 1
}

# Prompt for destination directory
$defaultDest = "C:\Archive\SalesPad\XML"
$destDir = Read-Host "Enter destination directory [$defaultDest]"
if ([string]::IsNullOrWhiteSpace($destDir)) {
    $destDir = $defaultDest
}

# Clean up path: trim whitespace and trailing backslash, remove surrounding quotes if user added them
$destDir = $destDir.Trim().Trim('"').Trim("'").TrimEnd('\')

# Prompt for file extension
$defaultExt = "*.xml"
$extension = Read-Host "Enter file extension to search for [$defaultExt]"
if ([string]::IsNullOrWhiteSpace($extension)) {
    $extension = $defaultExt
}

# Prompt for cutoff year
$yearInput = Read-Host "Enter year to archive BEFORE (files older than 1/1/YEAR), or leave blank for ALL files"
$cutoffYear = $null
$cutoffDate = $null
$maxAgeDate = $null

if (-not [string]::IsNullOrWhiteSpace($yearInput)) {
    $cutoffYear = [int]$yearInput
    # Calculate the cutoff date (January 1st of the specified year)
    $cutoffDate = Get-Date -Year $cutoffYear -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0
    # Format date for robocopy /MAXAGE parameter (YYYYMMDD)
    $maxAgeDate = $cutoffDate.ToString("yyyyMMdd")
}

# Build log file path (in script directory)
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($scriptDir)) {
    $scriptDir = Get-Location
}
$logFile = Join-Path -Path $scriptDir -ChildPath "Archive-OldFiles.log"

# Build robocopy arguments as a proper array (handles spaces automatically)
$robocopyArgs = @(
    $sourceDir
    $destDir
    $extension
)

# Only add date filter if a cutoff year was specified
if ($maxAgeDate) {
    $robocopyArgs += "/MINAGE:$maxAgeDate"
}

# Only include subdirectories if -Recurse is specified
if ($Recurse) {
    $robocopyArgs += "/E"
}

$robocopyArgs += @(
    "/R:3"
    "/W:5"
    "/NP"
    "/LOG+:$logFile"
)

Write-Host ""
Write-Host ("=" * 70) -ForegroundColor Cyan
if ($DryRun) {
    Write-Host "ARCHIVE OPERATION PREVIEW (DRY-RUN)" -ForegroundColor Magenta
} else {
    Write-Host "ARCHIVE OPERATION PREVIEW" -ForegroundColor Cyan
}
Write-Host ("=" * 70) -ForegroundColor Cyan
Write-Host ""
Write-Host "Source:      $sourceDir" -ForegroundColor Yellow
Write-Host "Destination: $destDir" -ForegroundColor Yellow
Write-Host "Extension:   $extension" -ForegroundColor Yellow
if ($cutoffDate) {
    Write-Host "Cutoff Date: Files older than $($cutoffDate.ToString('MM/dd/yyyy'))" -ForegroundColor Yellow
} else {
    Write-Host "Cutoff Date: ALL files (no date filter)" -ForegroundColor Yellow
}
if ($Recurse) {
    Write-Host "Subfolders:  Yes (recursive)" -ForegroundColor Yellow
} else {
    Write-Host "Subfolders:  No (root folder only)" -ForegroundColor Yellow
}
Write-Host ""

# Find matching files for preview using Get-ChildItem with -LiteralPath
Write-Host "Searching for matching files..." -ForegroundColor Gray
$getChildItemParams = @{
    LiteralPath = $sourceDir
    Filter = $extension
    File = $true
    ErrorAction = 'SilentlyContinue'
}
if ($Recurse) {
    $getChildItemParams.Recurse = $true
}

if ($cutoffDate) {
    $matchingFiles = Get-ChildItem @getChildItemParams |
        Where-Object { $_.LastWriteTime -lt $cutoffDate } |
        Sort-Object LastWriteTime
} else {
    $matchingFiles = Get-ChildItem @getChildItemParams |
        Sort-Object LastWriteTime
}

$totalFiles = $matchingFiles.Count
$totalSize = ($matchingFiles | Measure-Object -Property Length -Sum).Sum
$totalSizeMB = [math]::Round($totalSize / 1MB, 2)

Write-Host ""
Write-Host "FILES TO BE ARCHIVED:" -ForegroundColor Green
Write-Host ("-" * 70) -ForegroundColor Green
Write-Host "Total files found: $totalFiles" -ForegroundColor White
Write-Host "Total size: $totalSizeMB MB" -ForegroundColor White
Write-Host ""

if ($totalFiles -eq 0) {
    Write-Host "No files match the specified criteria." -ForegroundColor Yellow
    exit 0
}

# Show first 10 files
Write-Host "First 10 files (sorted by date, oldest first):" -ForegroundColor Cyan
Write-Host ""
$matchingFiles | Select-Object -First 10 | ForEach-Object {
    Write-Host "  $($_.LastWriteTime.ToString('yyyy-MM-dd'))  $($_.FullName)" -ForegroundColor Gray
}

if ($totalFiles -gt 10) {
    Write-Host "  ... and $($totalFiles - 10) more files" -ForegroundColor DarkGray
}

# Build command preview string with proper quoting for display
$commandPreview = "robocopy `"$sourceDir`" `"$destDir`" $extension"
if ($maxAgeDate) {
    $commandPreview += " /MINAGE:$maxAgeDate"
}
if ($Recurse) {
    $commandPreview += " /E"
}
$commandPreview += " /R:3 /W:5 /NP /LOG+:`"$logFile`""

Write-Host ""
Write-Host "COMMAND TO EXECUTE:" -ForegroundColor Cyan
Write-Host ("-" * 70) -ForegroundColor Cyan
Write-Host $commandPreview -ForegroundColor White
Write-Host ""
Write-Host ("=" * 70) -ForegroundColor Cyan

# Handle dry-run mode
if ($DryRun) {
    Write-Host ""
    Write-Host "DRY-RUN MODE: No files were copied." -ForegroundColor Magenta
    Write-Host "Run without -DryRun parameter to execute the archive operation." -ForegroundColor Magenta
    exit 0
}

# Confirm execution
Write-Host ""
$confirm = Read-Host "Do you want to proceed with the archive operation? (Y/N)"

if ($confirm -eq 'Y' -or $confirm -eq 'y') {
    # Create destination directory if it doesn't exist
    if (-not (Test-Path -LiteralPath $destDir)) {
        Write-Host "Creating destination directory: $destDir" -ForegroundColor Yellow
        New-Item -Path $destDir -ItemType Directory -Force | Out-Null
    }

    Write-Host ""
    Write-Host "Executing archive operation..." -ForegroundColor Green
    Write-Host ""

    # Execute robocopy using call operator with argument array (handles spaces properly)
    & robocopy @robocopyArgs

    $exitCode = $LASTEXITCODE

    Write-Host ""
    Write-Host ("=" * 70) -ForegroundColor Cyan

    # Interpret robocopy exit codes
    switch ($exitCode) {
        0 { Write-Host "No files were copied. No failure was encountered." -ForegroundColor Yellow }
        1 { Write-Host "Files were copied successfully." -ForegroundColor Green }
        2 { Write-Host "Extra files or directories were detected." -ForegroundColor Yellow }
        3 { Write-Host "Files were copied. Extra files were detected." -ForegroundColor Green }
        4 { Write-Host "Mismatched files or directories were detected." -ForegroundColor Yellow }
        5 { Write-Host "Files were copied. Mismatched files were detected." -ForegroundColor Green }
        6 { Write-Host "Extra and mismatched files were detected." -ForegroundColor Yellow }
        7 { Write-Host "Files were copied. Extra and mismatched files were detected." -ForegroundColor Green }
        8 { Write-Host "Some files could not be copied." -ForegroundColor Red }
        16 { Write-Host "Serious error. No files were copied." -ForegroundColor Red }
        default { Write-Host "Robocopy completed with exit code: $exitCode" -ForegroundColor White }
    }

    Write-Host "Log file: $logFile" -ForegroundColor Cyan
    Write-Host ("=" * 70) -ForegroundColor Cyan
} else {
    Write-Host ""
    Write-Host "Operation cancelled by user." -ForegroundColor Yellow
}
