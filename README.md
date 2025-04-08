# IntuneWin Packaging Tool

This PowerShell script automates the process of packaging a folder into a `.intunewin` file for Microsoft Intune deployments. It uses the `IntuneWinAppUtil.exe` tool, downloading it automatically if necessary, and provides user-friendly error handling and output messages.
It automatically generates an App in Intune, with Icon, Detection Rule and relevant Metadata.

## Features

- Automatically downloads `IntuneWinAppUtil.exe` if not found in the script's directory.
- The required Powershell modules will be downloaded automatically to the CurrentUser scope, when not already installed.
- Packages the provided folder into a `.intunewin` file with proper naming.
- Supports drag-and-drop functionality for easy folder selection.
- Displays the full path of the generated `.intunewin` file upon success.
- If the parameter `-Upload` is set to `$true` (Default) it generates a Application in Intune.
- The Install.bat file is used as information source for the Metadata for Intune. Use the REM lines
  |Install.Bat       | Intune          |
  |------------------|-----------------|
  |REM DESCRIPTION   | Name            |
  |REM MANUFACTURER  | Publischer      |
  |REM LANGUAGE      | Not used        |
  |REM FILENAME      | Executable used for detection rule |
  |REM VERSION       | App Version     |
  |REM ASSETNUMBER   | Notes           |
  |REM OWNER         | Owner of the App|
  
- If you supply the executable file name for the installed Program and the Program is installed locally, a file detection rule is generated automatically. 
- If the Package uses MSI, a MSI detection Rule is always generated.
- The first Jpg/Jpeg or PNG file found in the source folder or subfolder will be used as Application Icon in Intune.
- If you want a special Icon to be used you will have to use the command line and supply the full file name.
- Automatically create an App registration and save the log-in data to a Settings.json.

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- Internet connection (for downloading `IntuneWinAppUtil.exe` if not already available)

## Usage

### Drag-and-Drop

1. Drag a folder onto the script file (`PackageIntune.ps1`).
2. The script will package the folder into a `.intunewin` file and save it in the `Output` directory.

### Command-Line

1. Open a PowerShell terminal.
2. Run the script with the `SourceDir` parameter:

   ```powershell
   .\PackageIntune.ps1 -SourceDir "C:\Path\To\Your\Folder"
   ```

## Script Output

- The packaged `.intunewin` file is saved in the `Output` directory within the script's folder.
- The script will display the full path of the generated file upon completion.
- In Intune an  App will be generated. (If Upload is set to $true)

## Error Handling

- If the `IntuneWinAppUtil.exe` tool is missing, the script will automatically download it.
- Errors during execution are highlighted.
  
## Example Output

```plaintext
==========================================
          IntuneWin Packaging Tool         
==========================================

Packaging with IntuneWinAppUtil.exe...
File successfully packaged as:
C:\Path\To\Output\YourFolderName.intunewin

```

## Notes

- Ensure that the `Install.bat` file exists in the source folder, as it is required by `IntuneWinAppUtil.exe`.
- The script automatically creates an `Output` directory in the same location as the script if it does not already exist.
- The Template folder is a template to create a package

## Troubleshooting

- **Missing Dependencies:** Ensure you have an active internet connection if the script needs to download `IntuneWinAppUtil.exe`.
- **File Not Found Errors:** Verify that the provided folder exists and contains the required files.

## License

This project is licensed under the MIT License. See the LICENSE file for details.

---

### Author

- **Thomas Hoins (Datagroup OIT)**  
- [GitHub Profile](https://github.com/ThomasHoins/Intunewin)

Feel free to contribute to this repository by submitting issues or pull requests!

