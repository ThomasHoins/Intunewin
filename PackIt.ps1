<#
.SYNOPSIS
    Packages a folder into a .intunewin file using IntuneWinAppUtil.exe.

.DESCRIPTION
    This script automates the packaging process for Intune deployments by using the IntuneWinAppUtil.exe tool. 
    It accepts a folder as input (via drag-and-drop or as an argument) and creates a corresponding .intunewin file.
    If the IntuneWinAppUtil.exe file is not found, it will be automatically downloaded from the official Microsoft repository.

.NOTES
    Version:        1.1
    Author:         Thomas Hoins (DATAGROUP OIT)
    Initial Date:   14.01.2025
    Changes:        14.01.2025 Added error handling, clean outputs, and timestamp-based renaming.
    Changes:        14.01.2025 Automatic download of IntuneWinAppUtil.exe, auto-exit after 10 seconds.

.LINK
    [Your Documentation or GitHub Link Here]

.PARAMETER SourceDir
    Specifies the folder to package. Can be provided via drag-and-drop or as a command-line argument.

.PARAMETER OutputDir
    Specifies the directory where the .intunewin file will be saved. Default is the "Output" folder in the script's directory.

.INPUTS
    Accepts a folder path as input, either via command-line arguments or drag-and-drop.

.OUTPUTS
    Creates a .intunewin file in the specified output directory.

.EXAMPLE
    Drag and drop a folder onto the script:
    PS> .\PackageIntune.ps1

    Use as a command-line tool:
    PS> .\PackageIntune.ps1 -SourceDir "C:\MyFolder"

#>

# Script Header
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "          IntuneWin Packaging Tool         " -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# Define paths
$outputDir = Join-Path -Path $PSScriptRoot -ChildPath "Output"
$intuneWinAppUtil = Join-Path -Path $PSScriptRoot -ChildPath "IntuneWinAppUtil.exe"
$installBat = "Install.bat"

# URL to download IntuneWinAppUtil.exe
$downloadUrl = "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/master/IntuneWinAppUtil.exe"

try {
    # Check if an argument was provided (Drag-and-Drop)
    if ($args.Count -eq 0) {
        throw "No source directory provided. Drag and drop a folder onto the script."
    }

    # Get the source directory from the dragged file/folder
    $sourceDir = $args[0]
    if (-not (Test-Path -Path $sourceDir)) {
        throw "The provided source directory does not exist: $sourceDir"
    }

    # Create Output directory silently
    if (-not (Test-Path -Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }

    # Check if IntuneWinAppUtil.exe exists, download if not
    if (-not (Test-Path -Path $intuneWinAppUtil)) {
        Write-Host "IntuneWinAppUtil.exe not found. Downloading..." -ForegroundColor Yellow
        try {
            Invoke-WebRequest -Uri $downloadUrl -OutFile $intuneWinAppUtil -ErrorAction Stop
            Write-Host "IntuneWinAppUtil.exe downloaded successfully." -ForegroundColor Green
        } catch {
            throw "Failed to download IntuneWinAppUtil.exe. Please check your internet connection or the download URL."
        }
    }

    # Run IntuneWinAppUtil.exe silently
    Write-Host "Packaging with IntuneWinAppUtil.exe..." -ForegroundColor Green
    & $intuneWinAppUtil -c $sourceDir -s $installBat -o $outputDir | Out-Null

    # Move and rename the generated file
    $generatedFile = Join-Path -Path $outputDir -ChildPath "Install.intunewin"
    if (-not (Test-Path -Path $generatedFile)) {
        throw "Generated file not found: $generatedFile"
    }

    # Preserve folder name, including dots, and rename the file
    $sourceFolderName = Split-Path -Leaf $sourceDir
    $renamedFile = Join-Path -Path $outputDir -ChildPath ("$sourceFolderName.intunewin")

    # Check if the renamed file already exists
    if (Test-Path -Path $renamedFile) {
        # Append timestamp to avoid overwriting
        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $renamedFile = Join-Path -Path $outputDir -ChildPath ("$sourceFolderName-$timestamp.intunewin")
        Write-Host "File already exists. Renaming to: $renamedFile" -ForegroundColor Yellow
    }

    Move-Item -Path $generatedFile -Destination $renamedFile -Force
    Write-Host "File successfully packaged as:" -ForegroundColor Green
    Write-Host $renamedFile -ForegroundColor Green

} catch {
    Write-Host ""
    Write-Host "==========================================" -ForegroundColor Red
    Write-Host "ERROR: An error occurred!" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host "==========================================" -ForegroundColor Red
    Read-Host "Press Enter to close the window"
    exit 1
}

Write-Host ""
Write-Host "==========================================" -ForegroundColor Green
Write-Host "Script completed successfully!" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green
Start-Sleep -Seconds 10  # Wait 10 seconds before closing
exit 0
