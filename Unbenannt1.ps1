Import-Module Microsoft.Graph.Devices.CorporateManagement

#Disconnect-MgGraph
#$TenantID = "22c3b957-8768-4139-8b5e-279747e3ecbf"
#$AppId = "3997b08b-ee9c-4528-9afd-dfccb3ef2535"
#$AppSecret = "u9D8Q~HX31tRrc-tPwojE02g8OvcP4VqSz5H2a7p"


#$SecureClientSecret = ConvertTo-SecureString -String $AppSecret -AsPlainText -Force
#$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AppId, $SecureClientSecret
# Connect to Microsoft Graph Using the Tenant ID and Client Secret Credential
#Connect-MgGraph -TenantId $TenantID -ClientSecretCredential $ClientSecretCredential

###Connect-MgGraph -Scopes "DeviceManagementApps.ReadWrite.All"
#https://github.com/microsoftgraph/microsoft-graph-docs-contrib/blob/main/api-reference/beta/resources/intune-apps-win32lobapp.md#json-representation
#https://learn.microsoft.com/de-de/graph/api/intune-apps-win32lobapp-create?view=graph-rest-1.0&tabs=powershell
Get-MgDeviceAppManagementMobileApp -MobileAppId 03d0ccd4-312c-4caf-b63e-6d66bae59aec |fl

$Sourcepath="C:\Intunewin\Don Ho_Notepad++_8.7.5_MUI"
$InstallFilePath="C:\Intunewin\Output\Don Ho_Notepad++_8.7.5_MUI.intunewin"
$ExecutableName="C:\Program Files\Notepad++\notepad++.exe"
$Iconpath="C:\Intunewin\Don Ho_Notepad++_8.7.5_MUI\Appicon.png"

$params = @{
    "@odata.type" = "microsoft.graph.win32LobApp"
    displayName = "Test App"
    publisher = "Test Publisher"
    installCommandLine = "install.cmd"
    uninstallCommandLine = "uninstall.cmd"
    applicableArchitectures = "x64"
    setupFilePath = $InstallFilePath
    fileName = "App.intunewin"
    largeIcon = @{
		"@odata.type" = "microsoft.graph.mimeContent"
		type = "Type value"
		value = [convert]::ToBase64String((get-content $Iconpath -encoding byte))
	}
    rules = @(
        @{
            "@odata.type" = "microsoft.graph.win32LobAppRegistryRule"
            ruleType = "detection"
            check32BitOn64System = $true
            keyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\TestPublisher\TestApp"
            valueName = "Installed"
            operationType = "exists" # Prüft, ob der Schlüssel existiert
        }
    )
	installExperience = @{
		"@odata.type" = "microsoft.graph.win32LobAppInstallExperience"
		runAsAccount = "user" #system, user
		deviceRestartBehavior = "basedOnReturnCode" #basedOnReturnCode, allow, suppress, force
	}

}

New-MgDeviceAppManagementMobileApp -BodyParameter $params -Verbose
#get-MgDeviceAppManagementMobileApp -MobileAppId 66e10262-54d7-461f-9906-22509b857851
#Set-MgDeviceAppManagementMobileApp -BodyParameter $params -Verbose
