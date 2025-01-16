function Get-IntuneWinXML{
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

Function Get-IntuneWinFile(){
    param(
    [Parameter(Mandatory=$true)]
    $SourceFile,
    [Parameter(Mandatory=$true)]
    $fileName
    )
    $Folder = "win32"
    $Directory = [System.IO.Path]::GetDirectoryName("$SourceFile")
    if(-Not (Test-Path "$Directory\$folder")){
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
        
function MakeRequest($verb, $collectionPath, $body){

	$uri = "$baseUrl$collectionPath";
	$request = "$verb $uri";
	
	$clonedHeaders = CloneObject $authToken;
	$clonedHeaders["content-length"] = $body.Length;
	$clonedHeaders["content-type"] = "application/json";

	if ($logRequestUris) { Write-Host $request; }
	if ($logHeaders) { WriteHeaders $clonedHeaders; }
	if ($logContent) { Write-Host -ForegroundColor Gray $body; }

	try
	{
		$response = Invoke-RestMethod $uri -Method $verb -Headers $clonedHeaders -Body $body;
		$response;
	}
	catch
	{
		Write-Host -ForegroundColor Red $request;
		Write-Host -ForegroundColor Red $_.Exception.Message;
		throw;
	}
}

function New-Win32Lob(){

    <#
    .SYNOPSIS
    This function is used to upload a Win32 Application to the Intune Service
    .DESCRIPTION
    This function is used to upload a Win32 Application to the Intune Service
    .EXAMPLE
    Upload-Win32Lob "C:\Packages\package.intunewin" -publisher "Microsoft" -description "Package"
    This example uses all parameters required to add an intunewin File into the Intune Service
    .NOTES
    NAME: New-Win32Lob
    #>
    
    [cmdletbinding()]
    
    param(
        [parameter(Mandatory=$true,Position=1)]
        [ValidateNotNullOrEmpty()]
        [string]$SourceFile,
    
        [parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$displayName,
    
        [parameter(Mandatory=$true,Position=2)]
        [ValidateNotNullOrEmpty()]
        [string]$publisher,
    
        [parameter(Mandatory=$true,Position=3)]
        [ValidateNotNullOrEmpty()]
        [string]$description,
    
        [parameter(Mandatory=$true,Position=4)]
        [ValidateNotNullOrEmpty()]
        $detectionRules,
    
        [parameter(Mandatory=$true,Position=5)]
        [ValidateNotNullOrEmpty()]
        $returnCodes,
    
        [parameter(Mandatory=$false,Position=6)]
        [ValidateNotNullOrEmpty()]
        [string]$installCmdLine,
    
        [parameter(Mandatory=$false,Position=7)]
        [ValidateNotNullOrEmpty()]
        [string]$uninstallCmdLine,
    
        [parameter(Mandatory=$false,Position=8)]
        [ValidateSet('system','user')]
        $installExperience = "system"
    )
    
    try	{

        $LOBType = "microsoft.graph.win32LobApp"
        Write-Host "Testing if SourceFile '$SourceFile' Path is valid..." -ForegroundColor Yellow
        Test-SourceFile "$SourceFile"
        Write-Host
        Write-Host "Creating JSON data to pass to the service..." -ForegroundColor Yellow

        # Funciton to read Win32LOB file
        $DetectionXML = Get-IntuneWinXML "$SourceFile"     
        $FileName = $DetectionXML.ApplicationInfo.FileName
        $SetupFileName = $DetectionXML.ApplicationInfo.SetupFile
        $mobileAppBody = @{
            "@odata.type" = "microsoft.graph.win32LobApp"
            displayName = $DisplayName
            publisher = $publisher
            displayVersion = $version
            description = $description
            installCommandLine = $installCmdLine
            uninstallCommandLine = $uninstallcmdline
            applicableArchitectures = "x64"
            minimumSupportedOperatingSystem = @{"v10_1607" = $true}
            setupFilePath = $SetupFileName
            fileName = $filename
            msiInformation = $null
            runAs32bit = $false
        }

        if($DetectionRules.'@odata.type' -contains "#microsoft.graph.win32LobAppPowerShellScriptDetection" -and @($DetectionRules).'@odata.type'.Count -gt 1){
            Write-Host
            Write-Warning "A Detection Rule can either be 'Manually configure detection rules' or 'Use a custom detection script'"
            Write-Warning "It can't include both..."
            Write-Host
            break
        }
        else {
            $mobileAppBody | Add-Member -MemberType NoteProperty -Name 'detectionRules' -Value $detectionRules
        }

        #ReturnCodes
        if($returnCodes){  
            $mobileAppBody | Add-Member -MemberType NoteProperty -Name 'returnCodes' -Value @($returnCodes)
        }
        else {
            Write-Host
            Write-Warning "Intunewin file requires ReturnCodes to be specified"
            Write-Warning "If you want to use the default ReturnCode run 'Get-DefaultReturnCodes'"
            Write-Host
            break
        }
        Write-Host
        Write-Host "Creating application in Intune..." -ForegroundColor Yellow
        $mobileApp = MakeRequest "POST" "mobileApps" ($mobileAppBody | ConvertTo-Json);

        # Get the content version for the new app (this will always be 1 until the new app is committed).
        Write-Host
        Write-Host "Creating Content Version in the service for the application..." -ForegroundColor Yellow
        $appId = $mobileApp.id;
        $contentVersionUri = "mobileApps/$appId/$LOBType/contentVersions";
        $contentVersion = MakeRequest "POST" $contentVersionUri "{}";

        # Encrypt file and Get File Information
        Write-Host
        Write-Host "Getting Encryption Information for '$SourceFile'..." -ForegroundColor Yellow

        $encryptionInfo = @{
            encryptionKey = $DetectionXML.ApplicationInfo.EncryptionInfo.EncryptionKey
            macKey = $DetectionXML.ApplicationInfo.EncryptionInfo.macKey
            initializationVector = $DetectionXML.ApplicationInfo.EncryptionInfo.initializationVector
            mac = $DetectionXML.ApplicationInfo.EncryptionInfo.mac
            profileIdentifier = "ProfileVersion1";
            fileDigest = $DetectionXML.ApplicationInfo.EncryptionInfo.fileDigest
            fileDigestAlgorithm = $DetectionXML.ApplicationInfo.EncryptionInfo.fileDigestAlgorithm
        }

        $fileEncryptionInfo = @{
            fileEncryptionInfo = $encryptionInfo
        }

        # Extracting encrypted file
        $IntuneWinFile = Get-IntuneWinFile "$SourceFile" -fileName "$filename"

        [int64]$Size = $DetectionXML.ApplicationInfo.UnencryptedContentSize
        $EncrySize = (Get-Item "$IntuneWinFile").Length

        # Create a new file for the app.
        Write-Host
        Write-Host "Creating a new file entry in Azure for the upload..." -ForegroundColor Yellow
        $contentVersionId = $contentVersion.id;
        $fileBody =  @{ 
            "@odata.type" = "#microsoft.graph.mobileAppContentFile"
            name = $filename
            size = $Size
            sizeEncrypted =  $EncrySize
            manifest = $null
            isDependency = $false
        }

        $filesUri = "mobileApps/$appId/$LOBType/contentVersions/$contentVersionId/files";
        $file = MakeRequest "POST" $filesUri ($fileBody | ConvertTo-Json);
    
        # Wait for the service to process the new file request.
        Write-Host
        Write-Host "Waiting for the file entry URI to be created..." -ForegroundColor Yellow
        $fileId = $file.id;
        $fileUri = "mobileApps/$appId/$LOBType/contentVersions/$contentVersionId/files/$fileId";
        $file = WaitForFileProcessing $fileUri "AzureStorageUriRequest";

        # Upload the content to Azure Storage.
        Write-Host
        Write-Host "Uploading file to Azure Storage..." -f Yellow

        $sasUri = $file.azureStorageUri;
        UploadFileToAzureStorage $file.azureStorageUri "$IntuneWinFile" $fileUri;

        # Need to Add removal of IntuneWin file
        $IntuneWinFolder = [System.IO.Path]::GetDirectoryName("$IntuneWinFile")
        Remove-Item "$IntuneWinFile" -Force

        # Commit the file.
        Write-Host
        Write-Host "Committing the file into Azure Storage..." -ForegroundColor Yellow
        $commitFileUri = "mobileApps/$appId/$LOBType/contentVersions/$contentVersionId/files/$fileId/commit";
        MakeRequest "POST" $commitFileUri ($fileEncryptionInfo | ConvertTo-Json);

        # Wait for the service to process the commit file request.
        Write-Host
        Write-Host "Waiting for the service to process the commit file request..." -ForegroundColor Yellow
        $file = WaitForFileProcessing $fileUri "CommitFile";

        # Commit the app.
        Write-Host
        Write-Host "Committing the file into Azure Storage..." -ForegroundColor Yellow
        $commitAppUri = "mobileApps/$appId";
        $commitAppBody = GetAppCommitBody $contentVersionId $LOBType;
        MakeRequest "PATCH" $commitAppUri ($commitAppBody | ConvertTo-Json);

        Write-Host "Sleeping for $sleep seconds to allow patch completion..." -f Magenta
        Start-Sleep $sleep
        Write-Host
    
    }
    
    catch {

        Write-Host "";
        Write-Host -ForegroundColor Red "Aborting with exception: $($_.Exception.ToString())";
    
    }
}
