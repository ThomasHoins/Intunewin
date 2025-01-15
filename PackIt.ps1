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
param (
    [Parameter(Mandatory = $false)]
    [string]$SourceDir = "C:\Intunewin\Don Ho_Notepad++_8.7.5_MUI",

    [Parameter(Mandatory = $false)]
    [string]$outputDir="C:\Intunewin\Output",

    [Parameter(Mandatory = $false)]
    [string]$InstallCmd="Install.bat"
)
If (-Not($OutputDir)){$OutputDir="$PSScriptRoot\Output"}

#------------------------ Functions ------------------------
#===========================================================

function New-IntuneWin32App {
    <#
    .SYNOPSIS
        This Function creates a new Intune Win32 App.

    .DESCRIPTION
        This takes the input path for the application and icon, 
        and tries th extract more information from the application folder to create
        a new Intune Win32 App.

    .PARAMETER AppPath
        Path to the application folder.

    .PARAMETER IconName
        Name of the icon file.

    .EXAMPLE
        New-IntuneWin32App -AppPath "Value1" -IconName "Value2"
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$AppPath,

        [Parameter(Mandatory = $true)]
        [string]$SourceDir,

        [Parameter(Mandatory = $false)]
        [string]$IconName="Appicon.png"
    )

# Function Code
$Iconpath = "$SourceDir\$IconName"
$installCmd = "install.bat"
$uninstallCmd = "uninstall.bat"
$installCmdString= get-content "$SourceDir\$installCmd"
$displayName = ($installCmdString -match "DESCRIPTION").Replace("REM DESCRIPTION","").Trim()
$publisher = ($installCmdString -match "MANUFACTURER").Replace("REM MANUFACTURER","").Trim()
$fileName = ($installCmdString -match "FILENAME").Replace("REM FILENAME","").Trim()
$version = ($installCmdString -match "VERSION").Replace("REM VERSION","").Trim()
$Icon= [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes("$($Iconpath)"))
$Description = $(get-childitem $SourceDir -Filter "Description*" | get-content -Encoding UTF8 |Out-String)
#$size= ((get-item( $AppPath)).Length).ToString()
If(($installCmdString -match "msiexec").Count -gt 0){
    $MSIName = (get-childitem $SourceDir -Filter "*.msi")[0].FullName
    $MSIProductCode = (Get-AppLockerFileInformation $MSIName |select -ExpandProperty Publisher).BinaryName
    $Rule=@{
        "@odata.type"= "#microsoft.graph.win32LobAppProductCodeRule"
        ruleType= "detection"
        productCode= $MSIProductCode
        }
}
Else {
    $filePath = (Get-ChildItem -Path "C:\Program*"  -Recurse -ErrorAction SilentlyContinue -Include $fileName -Depth 3).FullName
    $Rule=@{
        "@odata.type"= "microsoft.graph.win32LobAppFileSystemRule"
        "ruleType"= "detection"
        "path"= (Split-Path -Path $filePath -Parent)
        "fileOrFolderName"= (Split-Path -Path $filePath -Leaf)
        "check32BitOn64System"= $true
        "operationType"= "version"
        "operator"= "greaterThanOrEqual"
        "comparisonValue"= $version
        }
}


$params = @{
    "@odata.type" = "microsoft.graph.win32LobApp"
    displayName = $displayName
    publisher = $publisher
    displayVersion = $version
    description = $Description
    installCommandLine = $installCmd
    uninstallCommandLine = $uninstallCmd
    applicableArchitectures = "x64"
    setupFilePath = $installCmd
    fileName = Split-Path($AppPath) -Leaf
    publishingState = "notPublished"
    msiInformation = $null
    runAs32bit = $false
	largeIcon = @{
		type = "image/png"
		value = $Icon
	}
    rules = @(
        $Rule
    )
	installExperience = @{
		"@odata.type" = "microsoft.graph.win32LobAppInstallExperience"
		runAsAccount = "system" #system, user
		deviceRestartBehavior = "basedOnReturnCode" #basedOnReturnCode, allow, suppress, force
	}

}
New-MgDeviceAppManagementMobileApp -BodyParameter $params
}


# Script Header
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "          IntuneWin Packaging Tool         " -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# Define paths

$intuneWinAppUtil = Join-Path -Path $PSScriptRoot -ChildPath "IntuneWinAppUtil.exe"

# URL to download IntuneWinAppUtil.exe
$downloadUrl = "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/master/IntuneWinAppUtil.exe"

try {
    # Get the source directory from the dragged file/folder
    if (-not (Test-Path -Path $sourceDir)) {
        throw "The provided source directory does not exist: $sourceDir, Drag and drop a folder onto the script."
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

    
    # Preserve folder name, including dots, and rename the file
    $sourceFolderName = Split-Path -Leaf $sourceDir
    $renamedFile = "$outputDir\$($sourceFolderName).intunewin"
    
    # If the intunewin does not already exist, make it
    if (-not (Test-Path -Path $renamedFile)) {
        # Run IntuneWinAppUtil.exe silently
        Write-Host "Packaging with IntuneWinAppUtil.exe..." -ForegroundColor Green
        $null = & $intuneWinAppUtil -c $sourceDir -s $installCmd -o $outputDir 

        # Move and rename the generated file
        $generatedFile = "$outputDir\Install.intunewin"
        if (-not (Test-Path -Path $generatedFile)) {
            throw "Generated file not found: $generatedFile"
        }
        Move-Item -Path $generatedFile -Destination $renamedFile -Force
        Write-Host "File successfully packaged as:" -ForegroundColor Green
        Write-Host $renamedFile -ForegroundColor Green
    }
    else {
        Write-Host "File already exists. Renaming to: $renamedFile" -ForegroundColor Yellow
    }

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
Write-Host "intunewin generated successfully!" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green


New-IntuneWin32App -AppPath $renamedFile -SourceDir $sourceDir
exit 0
