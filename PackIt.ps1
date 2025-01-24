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
    https://github.com/tabs-not-spaces/Intune-App-Deploy/blob/master/tasks/Deploy.Functions.ps1
    https://learn.microsoft.com/en-us/rest/api/storageservices/put-blob?tabs=microsoft-entra-id
    https://stackoverflow.com/questions/69031080/using-only-a-sas-token-to-upload-in-powershell
    https://learn.microsoft.com/en-us/rest/api/storageservices/naming-and-referencing-containers--blobs--and-metadata
    https://learn.microsoft.com/de-de/troubleshoot/mem/intune/app-management/develop-deliver-working-win32-app-via-intune


    $azCopyUri = "https://aka.ms/downloadazcopy-v10-windows"

    Modules:
    Microsoft.Graph.Authentication 
    Microsoft.Graph.Devices.CorporateManagement
    #Az.Storage

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

function Wait-ForFileProcessing {
    [cmdletbinding()]
    param (
        $fileUri,
        $stage
    )
    $attempts = 600
    $successState = "$($stage)Success"
    $pendingState = "$($stage)Pending"
    $file = $null
    while ($attempts -gt 0) {
        $file = Invoke-MgGraphRequest -Method GET -Uri $fileUri
        if ($file.uploadState -eq $successState) {
            break
        }
        elseif ($file.uploadState -ne $pendingState) {
            Write-Host -ForegroundColor Red $_.Exception.Message
            throw "File upload state is not successful: $($file.uploadState)"
        }
        Start-Sleep 10
        $attempts--
    }
    if ($null -eq $file -or $file.uploadState -ne $successState) {
        throw "File request did not complete in the allotted time."
    }
    $file
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
    $displayName = ($installCmdString -match "REM DESCRIPTION").Replace("REM DESCRIPTION","").Trim()
    $publisher = ($installCmdString -match "REM MANUFACTURER").Replace("REM MANUFACTURER","").Trim()
    If ($installCmdString -match "REM FILENAME"){$fileName = ($installCmdString -match "REM FILENAME").Replace("REM FILENAME","").Trim()}
    $version = ($installCmdString -match "REM VERSION").Replace("REM VERSION","").Trim()

    $IntuneWinMetadata = Get-IntuneWinMetadata -FilePath $AppPath

    # Create the Win32 App in Intune if it does not exist
    $MobileAppID=(Get-MgDeviceAppManagementMobileApp | Where-Object {$_.DisplayName -eq $displayName}).Id
    If (-not $MobileAppID){
        $Icon = @{
            "@odata.type" = "microsoft.graph.mimeContent"
            type= "image/png"
            #value = [Convert]::ToBase64String((Get-Content -Path $Iconpath -Encoding Byte))
            value = [Convert]::ToBase64String((Get-Content -Path $Iconpath-AsByteStream -Raw))
            #Das klappt noch nicht
            }
        $Description = $(get-childitem $SourceDir -Filter "Description*" | get-content -Encoding UTF8 |Out-String)
        
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
            If($fileName){
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
            Else{
                $filePath =""
                Write-Host "No file path could be found. Please Update the file Rule manually!" -ForegroundColor Red
            }
            $Rule=@{
                "@odata.type"= "microsoft.graph.win32LobAppFileSystemRule"
                "ruleType"= "detection"
                "path"= ""
                "fileOrFolderName"= ""
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
            #displayVersion = $version
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
            largeIcon = $Icon
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
    
    $UploadFile = Get-intuneWinFile -SourceFile $AppPath -fileName $IntuneWinMetadata.ApplicationInfo.FileName
    #$UploadFile = $AppPath
    # Prepare File Upload to Azure Blob Storage
    $FileName = $IntuneWinMetadata.ApplicationInfo.FileName
    $Size = [int64]$IntuneWinMetadata.ApplicationInfo.UnencryptedContentSize
    $EncrySize = (Get-Item "$UploadFile").Length
    $fileBody =  @{ 
        "@odata.type" = "#microsoft.graph.mobileAppContentFile"
        name = $FileName
        size = $Size
        sizeEncrypted = $EncrySize
        manifest = $null
        isDependency = $false
    }
    # Get the Content file ID
    $fileUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$MobileAppID/microsoft.graph.win32LobApp/contentVersions/1/files"
    $file = Invoke-MgGraphRequest -Method POST -Uri $fileUri -Body ($fileBody | ConvertTo-Json)  
    $ContentFileId = $file.id

    $fileUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$MobileAppID/microsoft.graph.win32LobApp/contentVersions/1/files/$ContentFileId";
    $file = Wait-ForFileProcessing $fileUri "AzureStorageUriRequest"

    # Upload the file to Azure Blob Storage
    $AzBlobUri = $file.azureStorageUri
    $headers = @{
        "x-ms-blob-type" = "BlockBlob"
        "Content-Length" = $EncrySize
        "Content-Type" = "application/octet-stream"
        }
        
    #This did not work in any Version of PS
    #$result = Invoke-WebRequest -Method "PUT" -Uri $AzBlobUri -InFile $UploadFile -Headers $headers -Verbose -HttpVersion 2.0

    # Upload the file to Azure Blob Storage (this actually worked with PS7)
    [System.Uri]$uriObject = $AzBlobUri 
    $storageAccountName = $uriObject.DnsSafeHost.Split(".")[0]
    $sasToken = $uriObject.Query.Substring(1)
    $uploadPath = $uriObject.LocalPath.Substring(1)
    $container = $uploadPath.Split("/")[0]
    $blobPath = $uploadPath.Substring($container.Length+1,$uploadPath.Length - $container.Length-1)
    $storageContext = New-AzStorageContext -StorageAccountName $storageAccountName -SasToken $sasToken
    $blobUpload = Set-AzStorageBlobContent -File $UploadFile -Container $container -Context $storageContext -Blob $blobPath -Force
    Write-Host "Upload finished! Details: Name $($blobUpload.Name), ContentType $($blobUpload.ContentType), Length $($blobUpload.Length), LastModified $($blobUpload.LastModified)"


    #if($result.StatusCode -ne 201){
    #    throw "Failed to upload file to Azure Blob Storage. Status code: $($blobUpload.StatusCode)"
    #}

    # Commit the file
    $fileEncryptionInfo = @{    
        fileEncryptionInfo = @{
            encryptionKey = $IntuneWinMetadata.ApplicationInfo.EncryptionInfo.EncryptionKey
            macKey = $IntuneWinMetadata.ApplicationInfo.EncryptionInfo.macKey
            initializationVector = $IntuneWinMetadata.ApplicationInfo.EncryptionInfo.initializationVector
            mac = $IntuneWinMetadata.ApplicationInfo.EncryptionInfo.mac
            profileIdentifier = "ProfileVersion1";
            fileDigest = $IntuneWinMetadata.ApplicationInfo.EncryptionInfo.fileDigest
            fileDigestAlgorithm = $IntuneWinMetadata.ApplicationInfo.EncryptionInfo.fileDigestAlgorithm
        }
    }

    $commitFileUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$MobileAppID/microsoft.graph.win32LobApp/contentVersions/1/files/$ContentFileId/commit"
    $commitFileUri
    $result = Invoke-MgGraphRequest -Method POST $commitFileUri -Body ($fileEncryptionInfo |ConvertTo-Json)

    $file = Wait-ForFileProcessing $fileUri "commitFile"

    If($file.uploadState -ne "commitFileSuccess"){
        throw "Failed to commit file to Azure Blob Storage. Status code: $($file.uploadState)"
    }   

    # Commit the App
    $commitAppBody = @{ 
        "@odata.type" = "#microsoft.graph.win32LobApp"
        committedContentVersion = "1"
        }   
    $commitUri ="https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$MobileAppID"
    $result = Invoke-MgGraphRequest -Method PATCH $commitUri -Body ($commitAppBody | ConvertTo-Json) -ContentType 'application/json'
}

#------------------------ Main Script ------------------------

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
    #$renamedFile = "$outputDir\$($sourceFolderName).intunewin"
    $renamedFile = "$outputDir\Install.intunewin"
    
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
        #Move-Item -Path $generatedFile -Destination $renamedFile -Force
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
Connect-MgGraph -TenantId $TenantID -ClientSecretCredential $ClientSecretCredential -NoWelcome

New-IntuneWin32App -AppPath $renamedFile -SourceDir $sourceDir
exit 0
