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

    https://learn.microsoft.com/en-us/graph/api/intune-apps-win32lobapp-create?view=graph-rest-1.0&tabs=http
    https://github.com/microsoftgraph/powershell-intune-samples
    https://github.com/MSEndpointMgr/IntuneWin32App/blob/master/Public/Get-IntuneWin32AppMetaData.ps1
    https://github.com/microsoftgraph/powershell-intune-samples/blob/master/LOB_Application/Win32_Application_Add.ps1#L852
    https://ourcloudnetwork.com/how-to-use-invoke-mggraphrequest-with-powershell/
    https://developer.microsoft.com/en-us/graph/graph-explorer
    https://www.scriptinglibrary.com/languages/powershell/how-to-upload-files-to-azure-blob-storage-using-powershell-via-the-rest-api/

    Modules:
    Microsoft.Graph.Authentication 
    Microsoft.Graph.Devices.CorporateManagement

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
    [string]$SourceDir = "C:\Intunewin\keepassxc.org_KeePassXC_2.79_MUI",

    [Parameter(Mandatory = $false)]
    [string]$outputDir="C:\Intunewin\Output",

    [Parameter(Mandatory = $false)]
    [string]$InstallCmd="Install.bat"
)
If (-Not($OutputDir)){$OutputDir="$PSScriptRoot\Output"}

#------------------------ Functions ------------------------
#===========================================================


function Add-FileToAzureStorage{
    param(
    [Parameter(Mandatory=$true)]
    $sasUri,
    [Parameter(Mandatory=$true)]
    $filepath,
    [Parameter(Mandatory=$true)]
    $fileUri
    )
	try {
        $chunkSizeInBytes = 1024l * 1024l * 60
		# Start the timer for SAS URI renewal.
		$sasRenewalTimer = [System.Diagnostics.Stopwatch]::StartNew()
		# Find the file size and open the file.
		$fileSize = (Get-Item $filepath).length;
		$chunks = [Math]::Ceiling($fileSize / $chunkSizeInBytes);
		$reader = New-Object System.IO.BinaryReader([System.IO.File]::Open($filepath, [System.IO.FileMode]::Open));
		
		# Upload each chunk. Check whether a SAS URI renewal is required after each chunk is uploaded and renew if needed.
		$ids = @();
		for ($chunk = 0; $chunk -lt $chunks; $chunk++){

			$id = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($chunk.ToString("0000")))
			$ids += $id

			$start = $chunk * $chunkSizeInBytes;
			$length = [Math]::Min($chunkSizeInBytes, $fileSize - $start)
			$bytes = $reader.ReadBytes($length)
			
			$currentChunk = $chunk + 1		

            Write-Progress -Activity "Uploading File to Azure Storage" -status "Uploading chunk $currentChunk of $chunks" -percentComplete ($currentChunk / $chunks*100)

            $uri = "$sasUri&comp=block&blockid=$id"
            $iso = [System.Text.Encoding]::GetEncoding("iso-8859-1")
            $encodedBody = $iso.GetString($bytes)
            $headers = @{
                "x-ms-blob-type" = "BlockBlob"
            }

            return = Invoke-MgGraphRequest -Uri $uri -Method Put -Headers $headers -Body $encodedBody

			# Renew the SAS URI if 7 minutes have elapsed since the upload started or was renewed last.
			if ($currentChunk -lt $chunks -and $sasRenewalTimer.ElapsedMilliseconds -ge 450000){
                $renewalUri = "$fileUri/renewUpload"
                $actionBody = ""
                $null = Invoke-MgGraphRequest -Uri $renewalUri -Body $actionBody
                $sasRenewalTimer.Restart()
            }

		}
        Write-Progress -Completed -Activity "Uploading File to Azure Storage"
		$reader.Close();
	}

	finally {
		if ($null -ne $reader) { $reader.Dispose() }
    }
	
	# Finalize the upload.
    $uri = "$sasUri&comp=blocklist";
	$xml = '<?xml version="1.0" encoding="utf-8"?><BlockList>';
	foreach ($id in $ids){
		$xml += "<Latest>$id</Latest>";
	}
	$xml += '</BlockList>';
	try{
		Invoke-RestMethod $uri -Method Put -Body $xml;
	}
	catch{
		Write-Host -ForegroundColor Red $request;
		Write-Host -ForegroundColor Red $_.Exception.Message;
		throw;
	}
}


Function Get-IntuneWinFile{
    param(
    [Parameter(Mandatory=$true)]
    $SourceFile,
    [Parameter(Mandatory=$true)]
    $fileName
    )

    $Folder = "win32"
    $Directory = [System.IO.Path]::GetDirectoryName("$SourceFile")
    if(-not(Test-Path "$Directory\$folder")){
        New-Item -ItemType Directory -Path "$Directory" -Name "$folder" | Out-Null
    }

    Add-Type -Assembly System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead("$SourceFile")
    $zip.Entries | Where-Object {$_.Name -like "$filename" } | ForEach-Object {
        [System.IO.Compression.ZipFileExtensions]::ExtractToFile($_, "$Directory\$folder\$filename", $true)
        }
    $zip.Dispose()
    return "$Directory\$folder\$filename"
}

function Get-IntuneWinMetadata{
    <#
    .SYNOPSIS
        Retrieves meta data from the detection.xml file inside the packaged Win32 application .intunewin file.

    .DESCRIPTION
        Retrieves meta data from the detection.xml file inside the packaged Win32 application .intunewin file.

    .PARAMETER FilePath
        Specify an existing local path to where the Win32 app .intunewin file is located.

    .EXAMPLE
        Get-IntuneWinMetadata -FilePath "somePath\BlaBla.intunewin"
    .NOTES
        shortened version of 
        https://github.com/MSEndpointMgr/IntuneWin32App/blob/master/Public/Get-IntuneWin32AppMetaData.ps1
    #>

    param (
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )
    if (Test-Path -Path $FilePath) {
        # Check if file extension is intunewin
        if (([System.IO.Path]::GetExtension((Split-Path -Path $FilePath -Leaf))) -ne ".intunewin") {
            throw "Given file name '$(Split-Path -Path $FilePath -Leaf)'contains an unsupported file extension. Supported extension is '.intunewin'"
        }
    }
    else {
        throw "File or folder does not exist"
    }
    $null = Add-Type -AssemblyName "System.IO.Compression.FileSystem" -ErrorAction Stop -Verbose:$false
    try {
        $IntuneWin32AppFile = [System.IO.Compression.ZipFile]::OpenRead($FilePath)
        if ($null -ne $IntuneWin32AppFile) {
            $DetectionXMLFile = $IntuneWin32AppFile.Entries | Where-Object { $_.Name -like "detection.xml" }
            $FileStream = $DetectionXMLFile.Open()

            # Construct new stream reader, pass file stream and read XML content to the end of the file
            $StreamReader = New-Object -TypeName "System.IO.StreamReader" -ArgumentList $FileStream -ErrorAction Stop
            $DetectionXMLContent = [xml]($StreamReader.ReadToEnd())
            
            # Close and dispose objects to preserve memory usage
            $FileStream.Close()
            $StreamReader.Close()
            $IntuneWin32AppFile.Dispose()

            # Handle return value with XML content from detection.xml
            return $DetectionXMLContent
        }
    }
    catch [System.Exception] {
        Write-Warning -Message "An error occurred while reading application information from detection.xml file. Error message: $($_.Exception.Message)"
    }
}

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
    $displayName = ($installCmdString -match "DESCRIPTION").Replace("REM DESCRIPTION","")[0].Trim()
    $publisher = ($installCmdString -match "MANUFACTURER").Replace("REM MANUFACTURER","")[0].Trim()
    $fileName = ($installCmdString -match "FILENAME").Replace("REM FILENAME","")[0].Trim()
    $version = ($installCmdString -match "VERSION").Replace("REM VERSION","")[0].Trim()
    # [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes("$($Iconpath)"))


    $MobileAppID=(Get-MgDeviceAppManagementMobileApp | Where-Object {$_.DisplayName -like $displayName}).Id[0]
    If (-not $MobileAppID){
        $Icon = @{
            "@odata.type" = "microsoft.graph.mimeContent"
            type= "image/png"
            value= [System.IO.File]::ReadAllBytes($Iconpath)  
            }
        $Description = $(get-childitem $SourceDir -Filter "Description*" | get-content -Encoding UTF8 |Out-String)
        $IntuneWinMetadata = Get-IntuneWinMetadata -FilePath $AppPath
        If(($installCmdString -match "msiexec").Count -gt 0){
            $MSIName = (get-childitem $SourceDir -Filter "*.msi")[0].FullName
            $MSIProductCode = (Get-AppLockerFileInformation $MSIName |Select-Object -ExpandProperty Publisher).BinaryName
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
            setupFilePath = $IntuneWinMetadata.ApplicationInfo.SetupFile
            fileName = $IntuneWinMetadata.ApplicationInfo.FileName
            #size = [int]$IntuneWinMetadata.ApplicationInfo.UnencryptedContentSize
            publishingState = "notPublished"
            msiInformation = $null
            runAs32bit = $false
            largeIcon = @(
                $Icon
            )
            rules = @(
                $Rule
            )
            installExperience = @{
                "@odata.type" = "microsoft.graph.win32LobAppInstallExperience"
                runAsAccount = "system" #system, user
                deviceRestartBehavior = "basedOnReturnCode" #basedOnReturnCode, allow, suppress, force
            }
        }
        $MobileAppID = (New-MgDeviceAppManagementMobileApp -BodyParameter (ConvertTo-Json($params))).Id
    }
    $Size = [int64]$IntuneWinMetadata.ApplicationInfo.UnencryptedContentSize
    # Prepare File Upload to Azure Blob Storage


    $fileUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$MobileAppID/microsoft.graph.win32LobApp/contentVersions?$select=id";
    $ContentID = (Invoke-MgGraphRequest -Method GET -Uri $fileUri).value.id
    If($ContentID.Count -gt 1){
        $file = Get-MgDeviceAppManagementMobileAppAsWin32LobAppContentVersionFile -MobileAppId $MobileAppID -MobileAppContentId $ContentID
        $ContentFileId = $file.id
    }
    Else{
        $fileBody =  @{ 
            "@odata.type" = "#microsoft.graph.mobileAppContentFile"
            name = $AppPath
            size = $Size
            sizeEncrypted =  (Get-Item $AppPath).Length
            manifest = $null
            isDependency = $false
        }
        $file = New-MgDeviceAppManagementMobileAppAsWin32LobAppContentVersionFile -MobileAppId $MobileAppID -MobileAppContentId $ContentID -BodyParameter (ConvertTo-Json($fileBody))
        $ContentFileId = $file.id
    }   
    #$fileUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$MobileAppID/microsoft.graph.win32LobApp/contentVersions/1/files/$mobileAppContentFileId";
    #$file = Invoke-MgGraphRequest -Method GET -Uri $fileUri
    $file = (Get-MgDeviceAppManagementMobileAppAsWin32LobAppContentVersionFile -MobileAppId $MobileAppID -MobileAppContentId $ContentID -MobileAppContentFileId $ContentFileId)
    
    # Upload the file to Azure Blob Storage
    $AzBlobUri = $file.azureStorageUri
    #$IntuneWinFile = Get-IntuneWinFile "$SourceFile" -fileName "$filename"
    #UploadFileToAzureStorage $file.azureStorageUri "$IntuneWinFile" $fileUri;

   

    $HashArguments = @{
        uri = $AzBlobUri
        method = "Put"
        InputFilePath = $AppPath
        headers = @{"x-ms-blob-type" = "BlockBlob"}
    }
    $UploadStatus= Invoke-MgGraphRequest @HashArguments
    $fileUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$MobileAppID/microsoft.graph.win32LobApp/contentVersions/1/files/$ContentFileId"
    #Add-FileToAzureStorage -sasUri $AzBlobUri -filepath $AppPath -fileUri $fileUri

}

# Script Header
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "          IntuneWin Packaging Tool         " -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# Define paths

$intuneWinAppUtil = "$PSScriptRoot\IntuneWinAppUtil.exe"

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

Disconnect-MgGraph
$TenantID = "22c3b957-8768-4139-8b5e-279747e3ecbf"
$AppId = "3997b08b-ee9c-4528-9afd-dfccb3ef2535"
$AppSecret = "u9D8Q~HX31tRrc-tPwojE02g8OvcP4VqSz5H2a7p"
# Connect to Microsoft Graph Using the Tenant ID and Client Secret Credential
$SecureClientSecret = ConvertTo-SecureString -String $AppSecret -AsPlainText -Force
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AppId, $SecureClientSecret
Connect-MgGraph -TenantId $TenantID -ClientSecretCredential $ClientSecretCredential

New-IntuneWin32App -AppPath $renamedFile -SourceDir $sourceDir
exit 0
