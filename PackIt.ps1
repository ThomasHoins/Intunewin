<#
.SYNOPSIS
    Packages a folder into a .intunewin file using IntuneWinAppUtil.exe.

.DESCRIPTION
    This script automates the packaging process for Intune deployments by using the IntuneWinAppUtil.exe tool. 
    It accepts a folder as input (via drag-and-drop or as an argument) and creates a corresponding .intunewin file.
    If the IntuneWinAppUtil.exe file is not found, it will be automatically downloaded from the official Microsoft repository.

.NOTES
    Version:        2.8.0
    Author:         Thomas Hoins (DATAGROUP OIT)
    Initial Date:   14.01.2025
    Changes:        14.01.2025 Added error handling, clean outputs, and timestamp-based renaming.
    Changes:        14.01.2025 Automatic download of IntuneWinAppUtil.exe, auto-exit after 10 seconds.
    Changes:        15.01.2025 Added support for command-line arguments, improved error handling.
    Changes:        24.01.2025 Added support for PowerShell 7, added creation of inune app via graph.
    Changes:        24.01.2025 Added support for Icon detection, added support for Version info.
    Changes:        25.01.2025 Added notes and owner to the App.
    Changes:        25.01.2025 Added support for MSI detection.
    Changes:        26.01.2025 Added some Error handling and improved the output. Updated the documentation.
    Changes:        26.01.2025 Some more Bug fixes and improvements.
    Changes:        28.01.2025 Improved the way to connect to MGGraph, added the automatic creation of an App registration
    Changes:        28.01.2025 Improved the Errorhandling<#
.SYNOPSIS
    Packages a folder into a .intunewin file using IntuneWinAppUtil.exe.

.DESCRIPTION
    This script automates the packaging process for Intune deployments by using the IntuneWinAppUtil.exe tool. 
    It accepts a folder as input (via drag-and-drop or as an argument) and creates a corresponding .intunewin file.
    If the IntuneWinAppUtil.exe file is not found, it will be automatically downloaded from the official Microsoft repository.

.NOTES
    Version:        2.8.1
    Author:         Thomas Hoins (DATAGROUP OIT)
    Initial Date:   14.01.2025
    Changes:        14.01.2025 Added error handling, clean outputs, and timestamp-based renaming.
    Changes:        14.01.2025 Automatic download of IntuneWinAppUtil.exe, auto-exit after 10 seconds.
    Changes:        15.01.2025 Added support for command-line arguments, improved error handling.
    Changes:        24.01.2025 Added support for PowerShell 7, added creation of inune app via graph.
    Changes:        24.01.2025 Added support for Icon detection, added support for Version info.
    Changes:        25.01.2025 Added notes and owner to the App.
    Changes:        25.01.2025 Added support for MSI detection.
    Changes:        26.01.2025 Added some Error handling and improved the output. Updated the documentation.
    Changes:        26.01.2025 Some more Bug fixes and improvements.
    Changes:        28.01.2025 Improved the way to connect to MGGraph, added the automatic creation of an App registration
    Changes:        28.01.2025 Improved the Errorhandling
    Changes:        07.02.2025 Changed Permissions to minimal
    Changes:        13.02.2025 Changed the Connect-Intune function to make it more resilient, removed unused code
    Changes:        14.02.2025 Bug Fixes, we are adding a Dummy File if the intunewin is <9MB
    Changes:        14.02.2025 Create Shortcut to Drop On
    Changes:        04.03.2025 Minor Bug Fix
    Changes:        10.03.2025 Bug Fix icon and Decriptions ar now found in Subfolders.
    Changes:        05.03.2025 Added the ability to automatically install the App to find the detection rules
    Changes:        13.03.2025 Changed the way to detect tehe file version, Bug with description fixed, added a wait before closing the window
    Changes:        28.03.2025 Bugfix, wrong Install.bat, intunewin not created.
    Issues: 	Still having issues with the description, there is an issue with Special cahracters.

    

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
    https://blog.icewolf.ch/archive/2022/12/02/create-azure-ad-app-registration-with-microsoft-graph-powershell
    https://knowledge-junction.in/2024/05/06/msgraph-create-app-microsoft-entra/
    https://practical365.com/common-graph-api-errors-powershell/
    $azCopyUri = "https://aka.ms/downloadazcopy-v10-windows"

    Modules:
    Required Modules Az.Storage, Microsoft.Graph.Devices.CorporateManagement,Microsoft.Graph.Authentication will be installed if not present.

.LINK
    [Your Documentation or GitHub Link Here]

.PARAMETER SourceDir
    Specifies the folder to package. Can be provided via drag-and-drop or as a command-line argument.

.PARAMETER OutputDir
    Specifies the directory where the .intunewin file will be saved. Default is the "Output" folder in the script's directory.

.PARAMETER IconName
    Specifies the full path of the icon file to use. Default is "Appicon.png".
    The script will search for the icon file in the source directory and its subfolders. And use the first found *.png, *.jpg or *.jpeg file.

.PARAMETER AppPath
    Specifies the path to the application file.

.PARAMETER InstallCmd
    Specifies the name of the installation command file. Default is "install.bat".

.PARAMETER Upload
    Specifies whether to upload the generated .intunewin file to Intune. Default is $true.  

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
    [string]$SourceDir = "C:\intunewin\Microsoft_365_Visio_App_for_Enterprise_2408_0325_(x64)",

    [Parameter(Mandatory = $false)]
    [string]$outputDir="C:\Intunewin\Output",

    [Parameter(Mandatory = $false)]
    [bool]$Upload= $true,  

    [Parameter(Mandatory = $false)]
    [bool]$Install= $true,  
    
    [Parameter(Mandatory = $false)]
    [string]$IconName, 

    [Parameter(Mandatory = $false)]
    [string]$InstallCmd
)
# Fix for dropped on folders with spaces
If ($PSBoundParameters.ContainsKey('SourceDir')){
    $SourceDir = [string]$MyInvocation.BoundParameters.Values
    $OutputDir = "C:\Intunewin\Output"}
    $IntunewinDir = "C:\Intunewin"
If (-Not($OutputDir)){$OutputDir="$(Split-Path ($SourceDir))\Output"}


#------------------------ Functions ------------------------

Function Create-Shortcut{
    Param(
        [string]$TargetFile,
        [string]$Arguments,
        [string]$Iconpath,
        [int]$IconNumber,
        [string]$ShortcutFile,
        [string]$Workdir,
        [int]$Style #1 Normal, 3 Maximized, 7 Minimized
    )
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
    $Shortcut.TargetPath = $TargetFile
    If ($Iconpath) {$Shortcut.IconLocation = "$($Iconpath),$IconNumber"}
    $Shortcut.Arguments = "$Arguments"
    $Shortcut.WorkingDirectory = $Workdir
    $Shortcut.WindowStyle = $Style
    $Shortcut.Save()
} 

function Create-DummyFile{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$PackagDir,

        [Parameter(Mandatory = $true)]
        [int]$SizeMB

        )
    $sizeInBytes = $SizeMB * 1MB
    $randomData = [byte[]]::new($sizeInBytes)
    [System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($randomData)
    [System.IO.File]::WriteAllBytes("$PackagDir\dummyfile.txt", $randomData)
}

function Wait-ForFileProcessing {
    # Wait for the file to be processed we will check the file upload state every 10 seconds
    [cmdletbinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$fileUri,
        [Parameter(Mandatory = $true)]
        [string]$stage
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

function Get-IntuneWinFileAndMetadata {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FilePath   # Path to the .intunewin file to extract
    )

    # Error handling for invalid or missing file
    if (-not (Test-Path -Path $FilePath)) {
        throw "File does not exist: $FilePath"
    }

    # Error handling for unsupported file extension
    if (([System.IO.Path]::GetExtension((Split-Path -Path $FilePath -Leaf))) -ne ".intunewin") {
        throw "The file '$($FilePath)' does not have a supported extension. Only '.intunewin' files are supported."
    }

    # Initialize the extraction folder and file paths

    $Directory = [System.IO.Path]::GetDirectoryName($FilePath)
    $Folder = "win32"
    $ExtractedFilePath = ""

    # Create the folder if it does not exist
    try {
        if (-not (Test-Path "$Directory\$Folder")) {New-Item -ItemType Directory -Path "$Directory" -Name $Folder | Out-Null}
    }
    catch {
        throw "Error creating extraction folder '$Folder'. Error: $_"
    }

    # opening the .intunewin file as a ZIP
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [System.IO.Compression.ZipFile]::OpenRead($FilePath)
    }
    catch {
        throw "Error opening the file '$FilePath' as a ZIP archive. Error: $_"
    }

    # Extract the file 
    $FileName = "IntunePackage.intunewin"   # Name of the file to extract from the .intunewin archive
    try {
        $zip.Entries | Where-Object { $_.Name -like $FileName } | ForEach-Object {
            $ExtractedFilePath = "$Directory\$Folder\$FileName"
            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($_, $ExtractedFilePath, $true)
        }
    }
    catch {
        throw "Error extracting the file '$FileName' from the archive '$FilePath'. Error: $_"
    }

    # Error handling for reading detection.xml content (or other XML file)
    $FileName = "detection.xml"
    $DetectionXMLContent = $null
    try {
        $DetectionXMLFile = $zip.Entries | Where-Object { $_.Name -like "detection.xml" }
        if ($DetectionXMLFile) {
            $FileStream = $DetectionXMLFile.Open()
            
            # Construct a StreamReader to read the XML content
            $StreamReader = New-Object -TypeName "System.IO.StreamReader" -ArgumentList $FileStream
            $DetectionXMLContent = [xml]($StreamReader.ReadToEnd())

            # Close the streams
            $FileStream.Close()
            $StreamReader.Close()
        }
        else {
            throw "The file 'detection.xml' was not found in the archive '$FilePath'."
        }
    }
    catch {
        throw "Error reading 'detection.xml' content from the archive '$FilePath'. Error: $_"
    }

    # Dispose the zip object to free up resources
    try {
        $zip.Dispose()
    }
    catch {
        Write-Warning "Error disposing of the zip object. Error: $_"
    }

    # Return both extracted file path and XML content (if any)
    return @{ 
        ExtractedFile = $ExtractedFilePath
        DetectionXMLContent = $DetectionXMLContent
    }
}

function New-IntuneWin32App {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$AppPath,

        [Parameter(Mandatory = $false)]
        [bool]$Install= $true,

        [Parameter(Mandatory = $true)]
        [string]$SourceDir,

        [Parameter(Mandatory = $false)]
        [string]$IconName
    )

    # Check if the required modules are installed
    $modules = 'Az.Storage', 'Microsoft.Graph.Devices.CorporateManagement', 'Microsoft.Graph.Authentication', 'Microsoft.Graph.Applications'
    $installed = @((Get-Module $modules -ListAvailable).Name | Select-Object -Unique)
    $notInstalled = Compare-Object $modules $installed -PassThru

    # At least one module is missing.
    # Install the missing modules now.
    if ($notInstalled) { 
        Write-Host "Installing required modules..." -ForegroundColor Yellow
        Install-Module -Scope CurrentUser $notInstalled -Force -AllowClobber
    }

    # Connect to Microsoft Graph Using the Tenant ID and Client Secret Credential
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
    $null = Disconnect-MgGraph -ErrorAction SilentlyContinue
    Connect-Intune -SecretFile "$PSScriptRoot\appreg-intune-CreateIntuneApp-Script-ReadWrite-Prod.json" -AppName "appreg-intune-CreateIntuneApp-Script-ReadWrite" -ApplicationPermissions "DeviceManagementApps.ReadWrite.All" -Scopes "Application.ReadWrite.All"

    # Get the Metadata from the install.bat
    $installCmd = "install.bat"
    $uninstallCmd = "uninstall.bat"
    $installCmdString= get-content "$SourceDir\$installCmd"
    $displayName = ($installCmdString -match "REM DESCRIPTION").Replace("REM DESCRIPTION","").Trim()
    $publisher = ($installCmdString -match "REM MANUFACTURER").Replace("REM MANUFACTURER","").Trim()
    If ($installCmdString -match "REM FILENAME"){$fileName = ($installCmdString -match "REM FILENAME").Replace("REM FILENAME","").Trim()}
    $version = ($installCmdString -match "REM VERSION").Replace("REM VERSION","").Trim()
    If ($installCmdString -match "REM OWNER"){$owner = ($installCmdString -match "REM OWNER").Replace("REM OWNER","").Trim()}
    If ($installCmdString -match "REM ASSETNUMBER"){$notes = ($installCmdString -match "REM ASSETNUMBER").Replace("REM ASSETNUMBER","").Trim()}
    $IntuneWinData = Get-IntuneWinFileAndMetadata -FilePath $AppPath
    $IntuneWinMetadata = $IntuneWinData.DetectionXMLContent

    # Create the Win32 App in Intune if it does not exist
    $MobileAppID=(Get-MgDeviceAppManagementMobileApp | Where-Object {$_.DisplayName -eq $displayName}).Id

    If (-not $MobileAppID){
        # If no icon is supplied Search for the Icon
        If ([string]::IsNullOrEmpty($IconName)){
            Write-Host "Searching for Icon..." -ForegroundColor Yellow
            $Iconpath=(Get-childitem -Path $SourceDir -Include *.png,*.jpg,*jpeg -Recurse -Depth 1| Select-Object -First 1).FullName
        }
        If ($Iconpath -like "*.jpg" -or $Iconpath -like "*.jpeg"){
            $IconType = "image/jpeg"
            Write-Host "Icon found: $Iconpath" -ForegroundColor Green
        }
        elseif ($Iconpath -like "*.png"){
            $IconType = "image/png"
            Write-Host "Icon found: $Iconpath" -ForegroundColor Green
        }
        else {
            $Iconpath = ""
        }
        If ([string]::IsNullOrEmpty($Iconpath)){
            Write-Host "No Icon found. Please Update the Icon manually!" -ForegroundColor Red
        }

        If ($Iconpath){
            If ($PSVersionTable.PSVersion.Major -lt 7){
                $ImageValue = [Convert]::ToBase64String((Get-Content -Path $Iconpath -Encoding Byte))
            }
            else {
                $ImageValue = [Convert]::ToBase64String((Get-Content -Path $Iconpath -AsByteStream -Raw))
            }
            $Icon = @{
                "@odata.type" = "microsoft.graph.mimeContent"
                type = $IconType
                value =  $ImageValue
                }
            }
        Else {
            $Icon = $null
        }
        $Text= $Descr = ""
        $Description = $(get-childitem $SourceDir -Filter "Description*" -Recurse -Depth 1)
        If (-Not $Description){
            $Description = "No Description found. Please Update the Description manually!"
        }
        elseif ($Description.Count -gt 1){ 
            ForEach($File in $Description){
                $Text = Get-Content -Path $File.FullName -Encoding UTF8 -Raw
                $Descr += $Text + "`r`n`r`n`r`n"

            }
            $DescriptionText = [string]$Descr.TrimEnd("`r`n`r`n")
        }
        else {
            $DescriptionText = [string](Get-Content -Path $Description.FullName -Encoding UTF8 -Raw)
        }

        If(($installCmdString -match "msiexec").Count -gt 0){
            $MSIName = (get-childitem $SourceDir -Filter "*.msi" -Recurse -Depth 1)[0].FullName
            $MSIProductCode = (Get-AppLockerFileInformation $MSIName |Select-Object -ExpandProperty Publisher).BinaryName
            $Rule=@{
                "@odata.type"= "#microsoft.graph.win32LobAppProductCodeRule"
                ruleType= "detection"
                productCode= $MSIProductCode
                }
        }
        Else {
            If($fileName){
                Write-Host "Searching for File Path..." -ForegroundColor Yellow
                $filePath = (Get-ChildItem -Path "C:\Program*"  -Recurse -ErrorAction SilentlyContinue -Include $fileName -Depth 3).FullName
                If ($filePath){
                    $path= (Split-Path -Path $filePath -Parent)
                    $fileOrFolderName= (Split-Path -Path $filePath -Leaf)
                    $FileVersion = (Get-Item $filePath).VersionInfo.FileVersion
                    If($FileVersion -ne $version){
                        Write-Host "File Version ($version) does not match installed Version ($FileVersion), Please fix this manually! " -ForegroundColor Red
                    }
                }
                Else{
                    Write-Host "No file path could be found.Installing Application..." -ForegroundColor Yellow
                    If ($Install){
                        $null = Start-Process -FilePath "$SourceDir\$installCmd" -Wait -passthru -Verb RunAs
                        $filePath = (Get-ChildItem -Path "C:\Program*"  -Recurse -ErrorAction SilentlyContinue -Include $fileName -Depth 3).FullName
                        If ($filePath){
                            $path= (Split-Path -Path $filePath -Parent)
                            $fileOrFolderName= (Split-Path -Path $filePath -Leaf)
                            $version = (Get-Item $filePath).VersionInfo.FileVersion
                        }
                        Write-Host "Removing Application..." -ForegroundColor Yellow
                        $null = Start-Process -FilePath "$SourceDir\$uninstallCmd" -Wait -passthru -Verb RunAs
                        Else{
                            Write-Host "No file path could be found. Please Update the file Rule manually!" -ForegroundColor Red
                        }
                    }
                }
            }
            Else{
                Write-Host "No file path could be found. Please Update the file Rule manually!" -ForegroundColor Red
            }
            $Rule=@{
                "@odata.type"= "microsoft.graph.win32LobAppFileSystemRule"
                "ruleType"= "detection"
                "path"= $path
                "fileOrFolderName"= $fileOrFolderName
                "check32BitOn64System"= $true
                "operationType"= "version"
                "operator"= "greaterThanOrEqual"
                "comparisonValue"= $version
            }
            If($fileName){
                Write-Host "==========================================" -ForegroundColor Green
                Write-Host "File Rule created: $($Rule |Out-String)" -ForegroundColor Green
                Write-Host "==========================================" -ForegroundColor Green
            }
        }
    
        $params = @{
            "@odata.type" = "microsoft.graph.win32LobApp"
            displayName = $displayName
            publisher = $publisher
            description = $DescriptionText
            notes = $notes
            owner = $owner
            installCommandLine = $installCmd
            uninstallCommandLine = $uninstallCmd
            applicableArchitectures = "x64"
            setupFilePath = $IntuneWinMetadata.ApplicationInfo.SetupFile
            fileName = $IntuneWinMetadata.ApplicationInfo.FileName
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
            returnCodes  = @(
                @{"returnCode" = 0;"type" = "success"}, `
                @{"returnCode" = 1707;"type" = "success"}, `
                @{"returnCode" = 3010;"type" = "softReboot"}, `
                @{"returnCode" = 1641;"type" = "hardReboot"}, `
                @{"returnCode" = 1618;"type" = "retry"}
            )
        }
        $MobileAppID = (New-MgDeviceAppManagementMobileApp -BodyParameter (ConvertTo-Json($params))).Id
        if ($MobileAppID ) {
            Write-Host  "App created successfully. App ID: $MobileAppID" -ForegroundColor Green
         }
        else {
            Write-Host "Failed to create the App. Please check the parameters and try again." -ForegroundColor Red
            Read-Host "Press Enter to close the window"
            Exit 1
        }
    }

    # Prepare File Upload to Azure Blob Storage
    $UploadFile =$IntuneWinData.ExtractedFile
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
    Try{
        $file = Invoke-MgGraphRequest -Method POST -Uri $fileUri -Body ($fileBody | ConvertTo-Json) 
        $ContentFileId = $file.id
    }
    Catch{
        $file = Invoke-MgGraphRequest -Method Get -Uri $fileUri
        If($file.value.isCommitted -eq "True"){
            Write-Host "This App is already committed. Please create a new App!" -ForegroundColor Green
            Read-Host "Press Enter to close the window"
            Exit 0
        }
        Write-Host "App already exists. Using existing App." -ForegroundColor Yellow
    }

    # Wait for the AzureStorageUriRequest to be processed
    Write-Host "Uploading file to Azure Blob Storage..." -ForegroundColor Yellow
    $fileUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$MobileAppID/microsoft.graph.win32LobApp/contentVersions/1/files/$ContentFileId";
    $file = Wait-ForFileProcessing $fileUri "AzureStorageUriRequest"
       
    # Upload the file to Azure Blob Storage 
    #  Get the SAS Token and Storage Account Name
    [System.Uri]$uriObject = $file.azureStorageUri
    $storageAccountName = $uriObject.DnsSafeHost.Split(".")[0]
    $sasToken = $uriObject.Query.Substring(1)
    $uploadPath = $uriObject.LocalPath.Substring(1)
    $container = $uploadPath.Split("/")[0]
    $blobPath = $uploadPath.Substring($container.Length+1,$uploadPath.Length - $container.Length-1)
    $storageContext = New-AzStorageContext -StorageAccountName $storageAccountName -SasToken $sasToken
    #  do the actual file to Azure Blob Storage
    $blobUpload = Set-AzStorageBlobContent -File $UploadFile -Container $container -Context $storageContext -Blob $blobPath -Force
    Write-Host "Upload finished! Details: Name $($blobUpload.Name), ContentType $($blobUpload.ContentType), Length $($blobUpload.Length), LastModified $($blobUpload.LastModified)" -ForegroundColor Green

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
    # Remove the file from the local storage
    Remove-Item -Path (Split-Path $UploadFile) -Force -Recurse -ErrorAction SilentlyContinue

    $commitFileUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$MobileAppID/microsoft.graph.win32LobApp/contentVersions/1/files/$ContentFileId/commit"
    try{
        Invoke-MgGraphRequest -Method POST $commitFileUri -Body ($fileEncryptionInfo |ConvertTo-Json)
    }
    catch{
        Write-Host "Failed to commit file to Azure Blob Storage. Status code: $($_.Exception.Message)" -ForegroundColor Red
        Read-Host "Press Enter to close the window"
        Exit 1
    }
    Write-Host "Waiting for the file to be committed..." -ForegroundColor Yellow

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
    try{
    Invoke-MgGraphRequest -Method PATCH $commitUri -Body ($commitAppBody | ConvertTo-Json) -ContentType 'application/json'
    }   
    catch{
        Write-Host "Failed to commit file to App. Status code: $($_.Exception.Message)" -ForegroundColor Red
        Read-Host "Press Enter to close the window"
        Exit 1
    }
    Write-Host "App successfully committed!" -ForegroundColor Green

    #Fix Version and Description
    $displayversionBody = @{
        "@odata.type" = "#microsoft.graph.win32LobApp"
        displayVersion = $version
   }
   $fileUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$MobileAppID"
   Try{
       $null = Invoke-MgGraphRequest -Method PATCH -Uri $fileUri -Body ($displayversionBody | ConvertTo-Json) 
   }
   Catch{
       Write-Host "Failed to update the display version. Status code: $($_.Exception.Message)" -ForegroundColor Red
       Read-Host "Press Enter to close the window"
       Exit 1
   }

    Write-Host ""
    Write-Host "==========================================" -ForegroundColor Green
    Write-Host "Intune App generated successfully!" -ForegroundColor Green
    write-host "App ID: $MobileAppID" -ForegroundColor Green
    Write-Host "Name: $displayName" -ForegroundColor Green
    Write-Host "Version: $version" -ForegroundColor Green
    Write-Host "==========================================" -ForegroundColor Green
}

function Connect-Intune{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$SecretFile = "$env:Temp\Settings.json",
		[Parameter(Mandatory = $false)]
        [string]$Scopes = "Application.ReadWrite.OwnedBy",
        [Parameter(Mandatory = $false)]
        [string]$AppName = "appreg-inune-BootMediaBuilder-Script-ReadWrite",
		[Parameter(Mandatory = $false)]
		[string[]]$ApplicationPermissions = "DeviceManagementServiceConfig.ReadWrite.All, Organization.Read.All",
		[Parameter(Mandatory = $false)]
		[string[]]$DelegationPermissions = ""

    )
    If (Test-Path -Path $SecretFile){
		Write-Host "Reading Settings file..." -ForegroundColor Yellow
		$SecretSettings = Get-Content -Path $SecretFile | ConvertFrom-Json
		$TenantID = $SecretSettings.TenantID
		$AppID = $SecretSettings.AppID
		$AppSecret = $SecretSettings.AppSecret
		Write-Host "Settings file read successfully." -ForegroundColor Green
		Write-Host "Using App Secret to connect to Tenant: $TenantID" -ForegroundColor Green
		$SecureClientSecret = ConvertTo-SecureString -String $AppSecret -AsPlainText -Force
		$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AppId, $SecureClientSecret
		$null = Connect-MgGraph -TenantId $TenantID -ClientSecretCredential $ClientSecretCredential -NoWelcome

    	#Test if Permissions are correct
		$actscopes = (Get-MgContext | Select-Object -ExpandProperty Scopes).Split(" ")
		$IncorrectScopes = ""
		$AppPerms = $ApplicationPermissions.Split(",").Trim()
		foreach ($AppPerm in $AppPerms) {
			if ($actscopes -notcontains $AppPerm) {
				$IncorrectScopes += $AppPerm -join ","
			}
		}
		if ($IncorrectScopes) {
			Write-Host "==========================================" -ForegroundColor Red
			Write-Host " The following permissions are missing:" -ForegroundColor Red
			Write-Host " $IncorrectScopes" -ForegroundColor Green
			Write-Host " Make sure to grant admin consent to your " -ForegroundColor Red
			Write-Host " API permissions in your newly created " -ForegroundColor Red
			Write-Host " App registration !!! " -ForegroundColor Red
			Write-Host "==========================================" -ForegroundColor Red
			Write-Host "https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade/quickStartType~/null/sourceType/Microsoft_AAD_IAM" -ForegroundColor Green
			Write-Host $Error[0].ErrorDetails
            Read-Host "Press Enter to close the window"
			Exit 1 
		}
		else{
			Write-Host "MS-Graph scopes: $($actscopes -join ", ") are correct" -ForegroundColor Green
		}
	}
	Else{
		Write-Host "Settings file not found. Creating a new one..." -ForegroundColor Yellow

		Connect-MgGraph -Scopes $Scopes -NoWelcome

		$TenantData =Get-MgContext
		$TenantID = $TenantData.TenantId

		#Create a new Application
		$AppObj = Get-MgApplication -Filter "DisplayName eq '$AppName'"
		If ($AppObj){
			$AppID = $AppObj.AppId
			Write-Host "App already exists. Updating existing App." -ForegroundColor Yellow
		}
		Else{
			Write-Host "Creating a new Application..." -ForegroundColor Yellow 
			$AppObj = New-MgApplication -DisplayName $AppName
			$AppID = $AppObj.AppId
			If($AppID){
				Write-Host "App created successfully. App ID: $AppID" -ForegroundColor Green
			}
			Else{
				Write-Host "Failed to create the App. Please check the parameters and try again." -ForegroundColor Red
                Read-Host "Press Enter to close the window"
				Exit 1  
			}
		}
		# Define Application and Delegation Permission ids and type in a hash
		$permissions = [ordered]@{}
		If ($ApplicationPermissions){
			$AppPermissions = $ApplicationPermissions.Split(",").Trim()
			$PermID = ""
			foreach($APermission in $AppPermissions){
				$PermID = (Find-MgGraphPermission $APermission -PermissionType Application -ExactMatch).Id
				$permissions.add($PermID,"Role")
			}
		}

		If ($DelegationPermissions){
			$DelPermissions = $DelegationPermissions.Split(",").Trim()
			$PermID = ""
			foreach($DPermission in $DelPermissions){
				$PermID = (Find-MgGraphPermission $DPermission -PermissionType Delegated -ExactMatch).Id
				$permissions.add($PermID,"Scope")
			}
		}

		# Build the accessBody for the hash
		$accessBody = [ordered]@{
			value = @(
				@{
					resourceAppId  = "00000003-0000-0000-c000-000000000000"
					resourceAccess = @()
				}
			)
		}

		# Add the  id/type pairs to the resourceAccess array
		foreach ($id in $permissions.Keys) {
			$accessBody.value[0].resourceAccess += @{
				id   = $id
				type = $permissions[$id]
			}
		}

		# Aplly upload the selected permissions via Graph API
		$fileUri = "https://graph.microsoft.com/v1.0/applications/$($AppObj.ID)/RequiredResourceAccess"
		try{
			$null = Invoke-MgGraphRequest -Method PATCH -Uri $fileUri -Body ($accessBody | ConvertTo-Json -Depth 4) 
		}
		catch{
			Write-Host "Failed to update the Required Resource Access. Status code: $($_.Exception.Message)" -ForegroundColor Red
            Read-Host "Press Enter to close the window"
			Exit 1
		}

		$passwordCred = @{
			"displayName" = "Secret-$($AppName)"
			"endDateTime" = (Get-Date).AddMonths(+12)
		}
		$ClientSecret = Add-MgApplicationPassword -ApplicationId  $AppObj.ID -PasswordCredential $passwordCred

		$AppSecret = $ClientSecret.SecretText
		If($AppSecret){
			Write-Host "App Secret ($AppSecret) created successfully." -ForegroundColor Green
		}
		Else{
			Write-Host "Failed to create the App Secret. Please check the parameters and try again." -ForegroundColor Red
			Read-Host "Press Enter to close the window"
            Exit 1
		}

		#Update Settings file with gathered information
		$SecretSettings = [ordered]@{
			Comment1 = "Make sure to keep this secret safe. This secret can be used to connect to your tenant!"
			Comment2 = "The following permissions are granted with this secret:"
			ApplicationPermissions = $ApplicationPermissions
			DelegationPermissions = $DelegationPermissions
			AppName = $AppObj.DisplayName
			CreatedBy = $TenantData.Account
			TenantID = $TenantID
			AppID = $AppID
			AppSecret = $AppSecret
		}
		Out-File -FilePath $SecretFile -InputObject ($SecretSettings | ConvertTo-Json)

		Write-Host ""
		Write-Host "==========================================================" -ForegroundColor Red
		Write-Host " A new App Registration ""$($AppObj.DisplayName)"" " -ForegroundColor Green
		Write-Host " has been created." -ForegroundColor Green
		Write-Host " Make sure to grant admin consent to your " -ForegroundColor Red
		Write-Host " API permissions in your newly created " -ForegroundColor Red
		Write-Host " App registration !!! " -ForegroundColor Red
		Write-Host  "==========================================================" -ForegroundColor Red
		Write-Host " Use this URL to grant consent:" -ForegroundColor Green
		Write-Host "https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade/quickStartType~/null/sourceType/Microsoft_AAD_IAM" -ForegroundColor Green
        Read-Host "Press Enter to close the window"		
        Exit 0
	}
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

# Check if link exists an create one
$LinkPath = "$PSScriptRoot\PackIt.lnk"
if (-not (Test-Path -Path $LinkPath)) {
    $TargetFile = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
    $Arguments =  "-Executionpolicy Bypass -command ""$PSScriptRoot\PackIt.ps1"""
    $Iconpath = "C:\Windows\System32\shell32.dll"
    Create-Shortcut -TargetFile $TargetFile -ShortcutFile $LinkPath -Arguments $Arguments -Iconpath $Iconpath -IconNumber 12 -Workdir $PSScriptRoot -Style 1
}

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

    # Rename the output file with the source folder name
    $sourceFolderName = Split-Path -Leaf $sourceDir
    $renamedFile = "$outputDir\$($sourceFolderName).intunewin"

    
    # If the intunewin does not already exist, make it
    if (-not (Test-Path -Path $renamedFile)) {
        # Run IntuneWinAppUtil.exe silently
        Write-Host "Packaging with IntuneWinAppUtil.exe..." -ForegroundColor Green
        $folderSize = (Get-ChildItem -Path $sourceDir -Recurse | Measure-Object -Property Length -Sum).Sum
        If ([string]::IsNullOrEmpty($InstallCmd)) {
            $InstallCmd= (Get-ChildItem $sourceDir -Include install*bat -Recurse | Select-Object -First 1 ).Name
            }
        If ($folderSize -lt 9437184) {
            $sizeMB = [int]((9437184 - $folderSize) / 1048576)
        Create-DummyFile -PackagDir $sourceDir -SizeMB $sizeMB
        }
        
        $null = &$intuneWinAppUtil -c $sourceDir -s $installCmd -o $outputDir 

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

# Upload the generated .intunewin file to Intune and generate a application.
If ($Upload){
    New-IntuneWin32App -AppPath $renamedFile -SourceDir $sourceDir -IconName $IconName -Install $Install
}
Read-Host "Press Enter to close the window"
exit 0

    Changes:        07.02.2025 Changed Permissions to minimal
    Changes:        13.02.2025 Changed the Connect-Intune function to make it more resilient, removed unused code
    Changes:        14.02.2025 Bug Fixes, we are adding a Dummy File if the intunewin is <9MB
    Changes:        14.02.2025 Create Shortcut to Drop On
    Changes:        04.03.2025 Minor Bug Fix
    Changes:        10.03.2025 Bug Fix icon and Decriptions ar now found in Subfolders.
    Changes:        05.03.2025 Added the ability to automatically install the App to find the detection rules
    Changes:        13.03.2025 Changed the way to detect tehe file version, Bug with description fixed, added a wait before closing the window
    Issues: 	Still having issues with the description, there is an issue with Special cahracters.

    

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
    https://blog.icewolf.ch/archive/2022/12/02/create-azure-ad-app-registration-with-microsoft-graph-powershell
    https://knowledge-junction.in/2024/05/06/msgraph-create-app-microsoft-entra/
    https://practical365.com/common-graph-api-errors-powershell/
    $azCopyUri = "https://aka.ms/downloadazcopy-v10-windows"

    Modules:
    Required Modules Az.Storage, Microsoft.Graph.Devices.CorporateManagement,Microsoft.Graph.Authentication will be installed if not present.

.LINK
    [Your Documentation or GitHub Link Here]

.PARAMETER SourceDir
    Specifies the folder to package. Can be provided via drag-and-drop or as a command-line argument.

.PARAMETER OutputDir
    Specifies the directory where the .intunewin file will be saved. Default is the "Output" folder in the script's directory.

.PARAMETER IconName
    Specifies the full path of the icon file to use. Default is "Appicon.png".
    The script will search for the icon file in the source directory and its subfolders. And use the first found *.png, *.jpg or *.jpeg file.

.PARAMETER AppPath
    Specifies the path to the application file.

.PARAMETER InstallCmd
    Specifies the name of the installation command file. Default is "install.bat".

.PARAMETER Upload
    Specifies whether to upload the generated .intunewin file to Intune. Default is $true.  

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
    [bool]$Upload= $true,  

    [Parameter(Mandatory = $false)]
    [bool]$Install= $true,  
    
    [Parameter(Mandatory = $false)]
    [string]$IconName, 

    [Parameter(Mandatory = $false)]
    [string]$InstallCmd="Install.bat"
)
# Fix for dropped on folders with spaces
If ($PSBoundParameters.ContainsKey('SourceDir')){
    $SourceDir = [string]$MyInvocation.BoundParameters.Values
    $OutputDir = "C:\Intunewin\Output"}
    $IntunewinDir = "C:\Intunewin"
If (-Not($OutputDir)){$OutputDir="$(Split-Path ($SourceDir))\Output"}


#------------------------ Functions ------------------------

Function Create-Shortcut{
    Param(
        [string]$TargetFile,
        [string]$Arguments,
        [string]$Iconpath,
        [int]$IconNumber,
        [string]$ShortcutFile,
        [string]$Workdir,
        [int]$Style #1 Normal, 3 Maximized, 7 Minimized
    )
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
    $Shortcut.TargetPath = $TargetFile
    If ($Iconpath) {$Shortcut.IconLocation = "$($Iconpath),$IconNumber"}
    $Shortcut.Arguments = "$Arguments"
    $Shortcut.WorkingDirectory = $Workdir
    $Shortcut.WindowStyle = $Style
    $Shortcut.Save()
} 

function Create-DummyFile{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$PackagDir,

        [Parameter(Mandatory = $true)]
        [int]$SizeMB

        )
    $sizeInBytes = $SizeMB * 1MB
    $randomData = [byte[]]::new($sizeInBytes)
    [System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($randomData)
    [System.IO.File]::WriteAllBytes("$PackagDir\dummyfile.txt", $randomData)
}

function Wait-ForFileProcessing {
    # Wait for the file to be processed we will check the file upload state every 10 seconds
    [cmdletbinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$fileUri,
        [Parameter(Mandatory = $true)]
        [string]$stage
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

function Get-IntuneWinFileAndMetadata {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FilePath   # Path to the .intunewin file to extract
    )

    # Error handling for invalid or missing file
    if (-not (Test-Path -Path $FilePath)) {
        throw "File does not exist: $FilePath"
    }

    # Error handling for unsupported file extension
    if (([System.IO.Path]::GetExtension((Split-Path -Path $FilePath -Leaf))) -ne ".intunewin") {
        throw "The file '$($FilePath)' does not have a supported extension. Only '.intunewin' files are supported."
    }

    # Initialize the extraction folder and file paths

    $Directory = [System.IO.Path]::GetDirectoryName($FilePath)
    $Folder = "win32"
    $ExtractedFilePath = ""

    # Create the folder if it does not exist
    try {
        if (-not (Test-Path "$Directory\$Folder")) {New-Item -ItemType Directory -Path "$Directory" -Name $Folder | Out-Null}
    }
    catch {
        throw "Error creating extraction folder '$Folder'. Error: $_"
    }

    # opening the .intunewin file as a ZIP
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [System.IO.Compression.ZipFile]::OpenRead($FilePath)
    }
    catch {
        throw "Error opening the file '$FilePath' as a ZIP archive. Error: $_"
    }

    # Extract the file 
    $FileName = "IntunePackage.intunewin"   # Name of the file to extract from the .intunewin archive
    try {
        $zip.Entries | Where-Object { $_.Name -like $FileName } | ForEach-Object {
            $ExtractedFilePath = "$Directory\$Folder\$FileName"
            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($_, $ExtractedFilePath, $true)
        }
    }
    catch {
        throw "Error extracting the file '$FileName' from the archive '$FilePath'. Error: $_"
    }

    # Error handling for reading detection.xml content (or other XML file)
    $FileName = "detection.xml"
    $DetectionXMLContent = $null
    try {
        $DetectionXMLFile = $zip.Entries | Where-Object { $_.Name -like "detection.xml" }
        if ($DetectionXMLFile) {
            $FileStream = $DetectionXMLFile.Open()
            
            # Construct a StreamReader to read the XML content
            $StreamReader = New-Object -TypeName "System.IO.StreamReader" -ArgumentList $FileStream
            $DetectionXMLContent = [xml]($StreamReader.ReadToEnd())

            # Close the streams
            $FileStream.Close()
            $StreamReader.Close()
        }
        else {
            throw "The file 'detection.xml' was not found in the archive '$FilePath'."
        }
    }
    catch {
        throw "Error reading 'detection.xml' content from the archive '$FilePath'. Error: $_"
    }

    # Dispose the zip object to free up resources
    try {
        $zip.Dispose()
    }
    catch {
        Write-Warning "Error disposing of the zip object. Error: $_"
    }

    # Return both extracted file path and XML content (if any)
    return @{ 
        ExtractedFile = $ExtractedFilePath
        DetectionXMLContent = $DetectionXMLContent
    }
}

function New-IntuneWin32App {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$AppPath,

        [Parameter(Mandatory = $false)]
        [bool]$Install= $true,

        [Parameter(Mandatory = $true)]
        [string]$SourceDir,

        [Parameter(Mandatory = $false)]
        [string]$IconName
    )

    # Check if the required modules are installed
    $modules = 'Az.Storage', 'Microsoft.Graph.Devices.CorporateManagement', 'Microsoft.Graph.Authentication', 'Microsoft.Graph.Applications'
    $installed = @((Get-Module $modules -ListAvailable).Name | Select-Object -Unique)
    $notInstalled = Compare-Object $modules $installed -PassThru

    # At least one module is missing.
    # Install the missing modules now.
    if ($notInstalled) { 
        Write-Host "Installing required modules..." -ForegroundColor Yellow
        Install-Module -Scope CurrentUser $notInstalled -Force -AllowClobber
    }

    # Connect to Microsoft Graph Using the Tenant ID and Client Secret Credential
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
    $null = Disconnect-MgGraph -ErrorAction SilentlyContinue
    Connect-Intune -SecretFile "$PSScriptRoot\appreg-intune-CreateIntuneApp-Script-ReadWrite-Prod.json" -AppName "appreg-intune-CreateIntuneApp-Script-ReadWrite" -ApplicationPermissions "DeviceManagementApps.ReadWrite.All" -Scopes "Application.ReadWrite.All"

    # Get the Metadata from the install.bat
    $installCmd = "install.bat"
    $uninstallCmd = "uninstall.bat"
    $installCmdString= get-content "$SourceDir\$installCmd"
    $displayName = ($installCmdString -match "REM DESCRIPTION").Replace("REM DESCRIPTION","").Trim()
    $publisher = ($installCmdString -match "REM MANUFACTURER").Replace("REM MANUFACTURER","").Trim()
    If ($installCmdString -match "REM FILENAME"){$fileName = ($installCmdString -match "REM FILENAME").Replace("REM FILENAME","").Trim()}
    $version = ($installCmdString -match "REM VERSION").Replace("REM VERSION","").Trim()
    If ($installCmdString -match "REM OWNER"){$owner = ($installCmdString -match "REM OWNER").Replace("REM OWNER","").Trim()}
    If ($installCmdString -match "REM ASSETNUMBER"){$notes = ($installCmdString -match "REM ASSETNUMBER").Replace("REM ASSETNUMBER","").Trim()}
    $IntuneWinData = Get-IntuneWinFileAndMetadata -FilePath $AppPath
    $IntuneWinMetadata = $IntuneWinData.DetectionXMLContent

    # Create the Win32 App in Intune if it does not exist
    $MobileAppID=(Get-MgDeviceAppManagementMobileApp | Where-Object {$_.DisplayName -eq $displayName}).Id

    If (-not $MobileAppID){
        # If no icon is supplied Search for the Icon
        If ([string]::IsNullOrEmpty($IconName)){
            Write-Host "Searching for Icon..." -ForegroundColor Yellow
            $Iconpath=(Get-childitem -Path $SourceDir -Include *.png,*.jpg,*jpeg -Recurse -Depth 1| Select-Object -First 1).FullName
        }
        If ($Iconpath -like "*.jpg" -or $Iconpath -like "*.jpeg"){
            $IconType = "image/jpeg"
            Write-Host "Icon found: $Iconpath" -ForegroundColor Green
        }
        elseif ($Iconpath -like "*.png"){
            $IconType = "image/png"
            Write-Host "Icon found: $Iconpath" -ForegroundColor Green
        }
        else {
            $Iconpath = ""
        }
        If ([string]::IsNullOrEmpty($Iconpath)){
            Write-Host "No Icon found. Please Update the Icon manually!" -ForegroundColor Red
        }

        If ($Iconpath){
            If ($PSVersionTable.PSVersion.Major -lt 7){
                $ImageValue = [Convert]::ToBase64String((Get-Content -Path $Iconpath -Encoding Byte))
            }
            else {
                $ImageValue = [Convert]::ToBase64String((Get-Content -Path $Iconpath -AsByteStream -Raw))
            }
            $Icon = @{
                "@odata.type" = "microsoft.graph.mimeContent"
                type = $IconType
                value =  $ImageValue
                }
            }
        Else {
            $Icon = $null
        }
        $Text= $Descr = ""
        $Description = $(get-childitem $SourceDir -Filter "Description*" -Recurse -Depth 1)
        If (-Not $Description){
            $Description = "No Description found. Please Update the Description manually!"
        }
        elseif ($Description.Count -gt 1){ 
            ForEach($File in $Description){
                $Text = Get-Content -Path $File.FullName -Encoding UTF8 -Raw
                $Descr += $Text + "`r`n`r`n`r`n"

            }
            $DescriptionText = [string]$Descr.TrimEnd("`r`n`r`n")
        }
        else {
            $DescriptionText = [string](Get-Content -Path $Description.FullName -Encoding UTF8 -Raw)
        }

        If(($installCmdString -match "msiexec").Count -gt 0){
            $MSIName = (get-childitem $SourceDir -Filter "*.msi" -Recurse -Depth 1)[0].FullName
            $MSIProductCode = (Get-AppLockerFileInformation $MSIName |Select-Object -ExpandProperty Publisher).BinaryName
            $Rule=@{
                "@odata.type"= "#microsoft.graph.win32LobAppProductCodeRule"
                ruleType= "detection"
                productCode= $MSIProductCode
                }
        }
        Else {
            If($fileName){
                Write-Host "Searching for File Path..." -ForegroundColor Yellow
                $filePath = (Get-ChildItem -Path "C:\Program*"  -Recurse -ErrorAction SilentlyContinue -Include $fileName -Depth 3).FullName
                If ($filePath){
                    $path= (Split-Path -Path $filePath -Parent)
                    $fileOrFolderName= (Split-Path -Path $filePath -Leaf)
                    $FileVersion = (Get-Item $filePath).VersionInfo.FileVersion
                    If($FileVersion -ne $version){
                        Write-Host "File Version ($version) does not match installed Version ($FileVersion), Please fix this manually! " -ForegroundColor Red
                    }
                }
                Else{
                    Write-Host "No file path could be found.Installing Application..." -ForegroundColor Yellow
                    If ($Install){
                        $null = Start-Process -FilePath "$SourceDir\$installCmd" -Wait -passthru -Verb RunAs
                        $filePath = (Get-ChildItem -Path "C:\Program*"  -Recurse -ErrorAction SilentlyContinue -Include $fileName -Depth 3).FullName
                        If ($filePath){
                            $path= (Split-Path -Path $filePath -Parent)
                            $fileOrFolderName= (Split-Path -Path $filePath -Leaf)
                            $version = (Get-Item $filePath).VersionInfo.FileVersion
                        }
                        Write-Host "Removing Application..." -ForegroundColor Yellow
                        $null = Start-Process -FilePath "$SourceDir\$uninstallCmd" -Wait -passthru -Verb RunAs
                        Else{
                            Write-Host "No file path could be found. Please Update the file Rule manually!" -ForegroundColor Red
                        }
                    }
                }
            }
            Else{
                Write-Host "No file path could be found. Please Update the file Rule manually!" -ForegroundColor Red
            }
            $Rule=@{
                "@odata.type"= "microsoft.graph.win32LobAppFileSystemRule"
                "ruleType"= "detection"
                "path"= $path
                "fileOrFolderName"= $fileOrFolderName
                "check32BitOn64System"= $true
                "operationType"= "version"
                "operator"= "greaterThanOrEqual"
                "comparisonValue"= $version
            }
            If($fileName){
                Write-Host "==========================================" -ForegroundColor Green
                Write-Host "File Rule created: $($Rule |Out-String)" -ForegroundColor Green
                Write-Host "==========================================" -ForegroundColor Green
            }
        }
    
        $params = @{
            "@odata.type" = "microsoft.graph.win32LobApp"
            displayName = $displayName
            publisher = $publisher
            description = $DescriptionText
            notes = $notes
            owner = $owner
            installCommandLine = $installCmd
            uninstallCommandLine = $uninstallCmd
            applicableArchitectures = "x64"
            setupFilePath = $IntuneWinMetadata.ApplicationInfo.SetupFile
            fileName = $IntuneWinMetadata.ApplicationInfo.FileName
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
            returnCodes  = @(
                @{"returnCode" = 0;"type" = "success"}, `
                @{"returnCode" = 1707;"type" = "success"}, `
                @{"returnCode" = 3010;"type" = "softReboot"}, `
                @{"returnCode" = 1641;"type" = "hardReboot"}, `
                @{"returnCode" = 1618;"type" = "retry"}
            )
        }
        $MobileAppID = (New-MgDeviceAppManagementMobileApp -BodyParameter (ConvertTo-Json($params))).Id
        if ($MobileAppID ) {
            Write-Host  "App created successfully. App ID: $MobileAppID" -ForegroundColor Green
         }
        else {
            Write-Host "Failed to create the App. Please check the parameters and try again." -ForegroundColor Red
            Read-Host "Press Enter to close the window"
            Exit 1
        }
    }

    # Prepare File Upload to Azure Blob Storage
    $UploadFile =$IntuneWinData.ExtractedFile
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
    Try{
        $file = Invoke-MgGraphRequest -Method POST -Uri $fileUri -Body ($fileBody | ConvertTo-Json) 
        $ContentFileId = $file.id
    }
    Catch{
        $file = Invoke-MgGraphRequest -Method Get -Uri $fileUri
        If($file.value.isCommitted -eq "True"){
            Write-Host "This App is already committed. Please create a new App!" -ForegroundColor Green
            Read-Host "Press Enter to close the window"
            Exit 0
        }
        Write-Host "App already exists. Using existing App." -ForegroundColor Yellow
    }

    # Wait for the AzureStorageUriRequest to be processed
    Write-Host "Uploading file to Azure Blob Storage..." -ForegroundColor Yellow
    $fileUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$MobileAppID/microsoft.graph.win32LobApp/contentVersions/1/files/$ContentFileId";
    $file = Wait-ForFileProcessing $fileUri "AzureStorageUriRequest"
       
    # Upload the file to Azure Blob Storage 
    #  Get the SAS Token and Storage Account Name
    [System.Uri]$uriObject = $file.azureStorageUri
    $storageAccountName = $uriObject.DnsSafeHost.Split(".")[0]
    $sasToken = $uriObject.Query.Substring(1)
    $uploadPath = $uriObject.LocalPath.Substring(1)
    $container = $uploadPath.Split("/")[0]
    $blobPath = $uploadPath.Substring($container.Length+1,$uploadPath.Length - $container.Length-1)
    $storageContext = New-AzStorageContext -StorageAccountName $storageAccountName -SasToken $sasToken
    #  do the actual file to Azure Blob Storage
    $blobUpload = Set-AzStorageBlobContent -File $UploadFile -Container $container -Context $storageContext -Blob $blobPath -Force
    Write-Host "Upload finished! Details: Name $($blobUpload.Name), ContentType $($blobUpload.ContentType), Length $($blobUpload.Length), LastModified $($blobUpload.LastModified)" -ForegroundColor Green

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
    # Remove the file from the local storage
    Remove-Item -Path (Split-Path $UploadFile) -Force -Recurse -ErrorAction SilentlyContinue

    $commitFileUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$MobileAppID/microsoft.graph.win32LobApp/contentVersions/1/files/$ContentFileId/commit"
    try{
        Invoke-MgGraphRequest -Method POST $commitFileUri -Body ($fileEncryptionInfo |ConvertTo-Json)
    }
    catch{
        Write-Host "Failed to commit file to Azure Blob Storage. Status code: $($_.Exception.Message)" -ForegroundColor Red
        Read-Host "Press Enter to close the window"
        Exit 1
    }
    Write-Host "Waiting for the file to be committed..." -ForegroundColor Yellow

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
    try{
    Invoke-MgGraphRequest -Method PATCH $commitUri -Body ($commitAppBody | ConvertTo-Json) -ContentType 'application/json'
    }   
    catch{
        Write-Host "Failed to commit file to App. Status code: $($_.Exception.Message)" -ForegroundColor Red
        Read-Host "Press Enter to close the window"
        Exit 1
    }
    Write-Host "App successfully committed!" -ForegroundColor Green

    #Fix Version and Description
    $displayversionBody = @{
        "@odata.type" = "#microsoft.graph.win32LobApp"
        displayVersion = $version
   }
   $fileUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$MobileAppID"
   Try{
       $null = Invoke-MgGraphRequest -Method PATCH -Uri $fileUri -Body ($displayversionBody | ConvertTo-Json) 
   }
   Catch{
       Write-Host "Failed to update the display version. Status code: $($_.Exception.Message)" -ForegroundColor Red
       Read-Host "Press Enter to close the window"
       Exit 1
   }

    Write-Host ""
    Write-Host "==========================================" -ForegroundColor Green
    Write-Host "Intune App generated successfully!" -ForegroundColor Green
    write-host "App ID: $MobileAppID" -ForegroundColor Green
    Write-Host "Name: $displayName" -ForegroundColor Green
    Write-Host "Version: $version" -ForegroundColor Green
    Write-Host "==========================================" -ForegroundColor Green
}

function Connect-Intune{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$SecretFile = "$env:Temp\Settings.json",
		[Parameter(Mandatory = $false)]
        [string]$Scopes = "Application.ReadWrite.OwnedBy",
        [Parameter(Mandatory = $false)]
        [string]$AppName = "appreg-inune-BootMediaBuilder-Script-ReadWrite",
		[Parameter(Mandatory = $false)]
		[string[]]$ApplicationPermissions = "DeviceManagementServiceConfig.ReadWrite.All, Organization.Read.All",
		[Parameter(Mandatory = $false)]
		[string[]]$DelegationPermissions = ""

    )
    If (Test-Path -Path $SecretFile){
		Write-Host "Reading Settings file..." -ForegroundColor Yellow
		$SecretSettings = Get-Content -Path $SecretFile | ConvertFrom-Json
		$TenantID = $SecretSettings.TenantID
		$AppID = $SecretSettings.AppID
		$AppSecret = $SecretSettings.AppSecret
		Write-Host "Settings file read successfully." -ForegroundColor Green
		Write-Host "Using App Secret to connect to Tenant: $TenantID" -ForegroundColor Green
		$SecureClientSecret = ConvertTo-SecureString -String $AppSecret -AsPlainText -Force
		$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AppId, $SecureClientSecret
		$null = Connect-MgGraph -TenantId $TenantID -ClientSecretCredential $ClientSecretCredential -NoWelcome

    	#Test if Permissions are correct
		$actscopes = (Get-MgContext | Select-Object -ExpandProperty Scopes).Split(" ")
		$IncorrectScopes = ""
		$AppPerms = $ApplicationPermissions.Split(",").Trim()
		foreach ($AppPerm in $AppPerms) {
			if ($actscopes -notcontains $AppPerm) {
				$IncorrectScopes += $AppPerm -join ","
			}
		}
		if ($IncorrectScopes) {
			Write-Host "==========================================" -ForegroundColor Red
			Write-Host " The following permissions are missing:" -ForegroundColor Red
			Write-Host " $IncorrectScopes" -ForegroundColor Green
			Write-Host " Make sure to grant admin consent to your " -ForegroundColor Red
			Write-Host " API permissions in your newly created " -ForegroundColor Red
			Write-Host " App registration !!! " -ForegroundColor Red
			Write-Host "==========================================" -ForegroundColor Red
			Write-Host "https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade/quickStartType~/null/sourceType/Microsoft_AAD_IAM" -ForegroundColor Green
			Write-Host $Error[0].ErrorDetails
            Read-Host "Press Enter to close the window"
			Exit 1 
		}
		else{
			Write-Host "MS-Graph scopes: $($actscopes -join ", ") are correct" -ForegroundColor Green
		}
	}
	Else{
		Write-Host "Settings file not found. Creating a new one..." -ForegroundColor Yellow

		Connect-MgGraph -Scopes $Scopes -NoWelcome

		$TenantData =Get-MgContext
		$TenantID = $TenantData.TenantId

		#Create a new Application
		$AppObj = Get-MgApplication -Filter "DisplayName eq '$AppName'"
		If ($AppObj){
			$AppID = $AppObj.AppId
			Write-Host "App already exists. Updating existing App." -ForegroundColor Yellow
		}
		Else{
			Write-Host "Creating a new Application..." -ForegroundColor Yellow 
			$AppObj = New-MgApplication -DisplayName $AppName
			$AppID = $AppObj.AppId
			If($AppID){
				Write-Host "App created successfully. App ID: $AppID" -ForegroundColor Green
			}
			Else{
				Write-Host "Failed to create the App. Please check the parameters and try again." -ForegroundColor Red
				Exit 1  
			}
		}
		# Define Application and Delegation Permission ids and type in a hash
		$permissions = [ordered]@{}
		If ($ApplicationPermissions){
			$AppPermissions = $ApplicationPermissions.Split(",").Trim()
			$PermID = ""
			foreach($APermission in $AppPermissions){
				$PermID = (Find-MgGraphPermission $APermission -PermissionType Application -ExactMatch).Id
				$permissions.add($PermID,"Role")
			}
		}

		If ($DelegationPermissions){
			$DelPermissions = $DelegationPermissions.Split(",").Trim()
			$PermID = ""
			foreach($DPermission in $DelPermissions){
				$PermID = (Find-MgGraphPermission $DPermission -PermissionType Delegated -ExactMatch).Id
				$permissions.add($PermID,"Scope")
			}
		}

		# Build the accessBody for the hash
		$accessBody = [ordered]@{
			value = @(
				@{
					resourceAppId  = "00000003-0000-0000-c000-000000000000"
					resourceAccess = @()
				}
			)
		}

		# Add the  id/type pairs to the resourceAccess array
		foreach ($id in $permissions.Keys) {
			$accessBody.value[0].resourceAccess += @{
				id   = $id
				type = $permissions[$id]
			}
		}

		# Aplly upload the selected permissions via Graph API
		$fileUri = "https://graph.microsoft.com/v1.0/applications/$($AppObj.ID)/RequiredResourceAccess"
		try{
			$null = Invoke-MgGraphRequest -Method PATCH -Uri $fileUri -Body ($accessBody | ConvertTo-Json -Depth 4) 
		}
		catch{
			Write-Host "Failed to update the Required Resource Access. Status code: $($_.Exception.Message)" -ForegroundColor Red
			Exit 1
		}

		$passwordCred = @{
			"displayName" = "Secret-$($AppName)"
			"endDateTime" = (Get-Date).AddMonths(+12)
		}
		$ClientSecret = Add-MgApplicationPassword -ApplicationId  $AppObj.ID -PasswordCredential $passwordCred

		$AppSecret = $ClientSecret.SecretText
		If($AppSecret){
			Write-Host "App Secret ($AppSecret) created successfully." -ForegroundColor Green
		}
		Else{
			Write-Host "Failed to create the App Secret. Please check the parameters and try again." -ForegroundColor Red
			Read-Host "Press Enter to close the window"
            Exit 1
		}

		#Update Settings file with gathered information
		$SecretSettings = [ordered]@{
			Comment1 = "Make sure to keep this secret safe. This secret can be used to connect to your tenant!"
			Comment2 = "The following permissions are granted with this secret:"
			ApplicationPermissions = $ApplicationPermissions
			DelegationPermissions = $DelegationPermissions
			AppName = $AppObj.DisplayName
			CreatedBy = $TenantData.Account
			TenantID = $TenantID
			AppID = $AppID
			AppSecret = $AppSecret
		}
		Out-File -FilePath $SecretFile -InputObject ($SecretSettings | ConvertTo-Json)

		Write-Host ""
		Write-Host "==========================================================" -ForegroundColor Red
		Write-Host " A new App Registration ""$($AppObj.DisplayName)"" " -ForegroundColor Green
		Write-Host " has been created." -ForegroundColor Green
		Write-Host " Make sure to grant admin consent to your " -ForegroundColor Red
		Write-Host " API permissions in your newly created " -ForegroundColor Red
		Write-Host " App registration !!! " -ForegroundColor Red
		Write-Host  "==========================================================" -ForegroundColor Red
		Write-Host " Use this URL to grant consent:" -ForegroundColor Green
		Write-Host "https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade/quickStartType~/null/sourceType/Microsoft_AAD_IAM" -ForegroundColor Green
        Read-Host "Press Enter to close the window"		
        Exit 0
	}
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

# Check if link exists an create one
$LinkPath = "$PSScriptRoot\PackIt.lnk"
if (-not (Test-Path -Path $LinkPath)) {
    $TargetFile = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
    $Arguments =  "-Executionpolicy Bypass -command ""$PSScriptRoot\PackIt.ps1"""
    $Iconpath = "C:\Windows\System32\shell32.dll"
    Create-Shortcut -TargetFile $TargetFile -ShortcutFile $LinkPath -Arguments $Arguments -Iconpath $Iconpath -IconNumber 12 -Workdir $PSScriptRoot -Style 1
}

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

    # Rename the output file with the source folder name
    $sourceFolderName = Split-Path -Leaf $sourceDir
    $renamedFile = "$outputDir\$($sourceFolderName).intunewin"

    
    # If the intunewin does not already exist, make it
    if (-not (Test-Path -Path $renamedFile)) {
        # Run IntuneWinAppUtil.exe silently
        Write-Host "Packaging with IntuneWinAppUtil.exe..." -ForegroundColor Green
        $folderSize = (Get-ChildItem -Path $sourceDir -Recurse | Measure-Object -Property Length -Sum).Sum
        If ($folderSize -lt 9437184) {
            $sizeMB = [int]((9437184 - $folderSize) / 1048576)
        Create-DummyFile -PackagDir $sourceDir -SizeMB $sizeMB
        }
        
        $null = &$intuneWinAppUtil -c $sourceDir -s $installCmd -o $outputDir 

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

# Upload the generated .intunewin file to Intune and generate a application.
If ($Upload){
    New-IntuneWin32App -AppPath $renamedFile -SourceDir $sourceDir -IconName $IconName -Install $Install
}
Read-Host "Press Enter to close the window"
exit 0
