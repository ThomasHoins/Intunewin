# PackIt change history

Version 2.11
Author: Thomas Hoins (DATAGROUP OIT)  
Initial date: 14.01.2025

## Changes

- 14.01.2025: Added error handling, clean outputs, and timestamp-based renaming.
- 14.01.2025: Automatic download of IntuneWinAppUtil.exe, auto-exit after 10 seconds.
- 15.01.2025: Added support for command-line arguments, improved error handling.
- 24.01.2025: Added support for PowerShell 7, added creation of inune app via graph.
- 24.01.2025: Added support for Icon detection, added support for Version info.
- 25.01.2025: Added notes and owner to the App.
- 25.01.2025: Added support for MSI detection.
- 26.01.2025: Added some Error handling and improved the output. Updated the documentation.
- 26.01.2025: Some more Bug fixes and improvements.
- 28.01.2025: Improved the way to connect to MGGraph, added the automatic creation of an App registration.
- 28.01.2025: Improved the Errorhandling.
- 07.02.2025: Changed Permissions to minimal.
- 13.02.2025: Changed the Connect-Intune function to make it more resilient, removed unused code.
- 14.02.2025: Bug Fixes, we are adding a Dummy File if the intunewin is <9MB.
- 14.02.2025: Create Shortcut to Drop On.
- 04.03.2025: Minor Bug Fix.
- 10.03.2025: Bug Fix icon and Decriptions ar now found in Subfolders.
- 05.03.2025: Added the ability to automatically install the App to find the detection rules.
- 13.03.2025: Changed the way to detect tehe file version, Bug with description fixed, added a wait before closing the window.
- 08.04.2025: Fixed a Bug, We did not consider already existing Versions and special characters in the install bat are fixed now.
- 18.09.2025: Added automatic install and uninstall command detection.
- 30.09.2025: Added automatic unlock for Internet files.
- 27.01.2026: Fixed MSI Rule detection and some other fixes.
- 13.02.2026: Fixed Installer detection.
- 20.02.2026: TEMP and TMP were temporarily redirected during IntuneWinAppUtil.exe execution to mitigate Cynet interference; original environment variables were restored after completion.
- 07.07.2026: Minor Bug Fixe regarding uninstallation for detection rule.
- 24.07.2026: Fixed version handling for Intune upload and ensured displayVersion uses dot notation.
- 24.07.2026: Added `-GenerateGroups` feature: creates Entra security groups from `groupTemplate.json` (install groups default to `available`; uninstall groups omit assignments).
- 24.07.2026: Added `-Supersedence` feature: automatically resolves and creates supersedence relationships for existing apps in Intune.
- 21.10.2026: Added `-AssignGroups` feature: assigns Entra security groups from `groupTemplate.json` (install groups default to `available`; uninstall groups omit assignments).
- 21.10.2026: Refactored Win32 app upload, metadata extraction, detection rule generation, asset loading, and payload construction into focused helper functions.
- 21.10.2026: Separated group creation and app assignment controls; existing groups are resolved by display name without a Group.Read.All pre-check.
- 21.10.2026: Improved app commit failure handling with concise Graph error details, app cleanup on failure, and preservation of extracted upload data until commit succeeds.
- 21.10.2026: Translated remaining German comments and messages to English; updated the script version to 2.10.3.
- 21.10.2026: Allowed detection rule generation to fall back when install command content is unavailable.
- 21.10.2026: Moved required module installation outside the app creation function and added optional module updates.
- 21.10.2026: Install only required modules that are missing or have a newer gallery version.
- 21.10.2026: Passed install and uninstall commands explicitly to the app upload workflow and validated the install script before connecting to Graph.

## Known issue

Still having issues with the description, there is an issue with Special characters.

## References

- https://learn.microsoft.com/en-us/graph/api/intune-apps-win32lobapp-create?view=graph-rest-1.0&tabs=http
- https://github.com/microsoftgraph/powershell-intune-samples
- https://github.com/MSEndpointMgr/IntuneWin32App/blob/master/Public/Get-IntuneWin32AppMetaData.ps1
- https://github.com/microsoftgraph/powershell-intune-samples/blob/master/LOB_Application/Win32_Application_Add.ps1#L852
- https://ourcloudnetwork.com/how-to-use-invoke-mggraphrequest-with-powershell/
- https://developer.microsoft.com/en-us/graph/graph-explorer
- https://www.scriptinglibrary.com/languages/powershell/how-to-upload-files-to-azure-blob-storage-using-powershell-via-the-rest-api/
- https://github.com/tabs-not-spaces/Intune-App-Deploy/blob/master/tasks/Deploy.Functions.ps1
- https://learn.microsoft.com/en-us/rest/api/storageservices/put-blob?tabs=microsoft-entra-id
- https://stackoverflow.com/questions/69031080/using-only-a-sas-token-to-upload-in-powershell
- https://learn.microsoft.com/en-us/rest/api/storageservices/naming-and-referencing-containers--blobs--and-metadata
- https://learn.microsoft.com/de-de/troubleshoot/mem/intune/app-management/develop-deliver-working-win32-app-via-intune
- https://blog.icewolf.ch/archive/2022/12/02/create-azure-ad-app-registration-with-microsoft-graph
- https://knowledge-junction.in/2024/05/06/msgraph-create-app-microsoft-entra/
- https://practical365.com/common-graph-api-errors-powershell/
- https://aka.ms/downloadazcopy-v10-windows

## Required modules

The required module list is maintained in `Install-RequiredModules`. Modules are installed and optionally updated according to the script's `-UpdateModules` parameter.
