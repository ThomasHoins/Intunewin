# IntuneWin Packaging Tool

This PowerShell script automates the process of packaging a folder into a `.intunewin` file for Microsoft Intune deployments. It uses the `IntuneWinAppUtil.exe` tool, downloading it automatically if necessary, and provides user-friendly error handling and output messages.

## Features

- Automatically downloads `IntuneWinAppUtil.exe` if not found in the script's directory.
- Packages the provided folder into a `.intunewin` file with proper naming.
- Supports drag-and-drop functionality for easy folder selection.
- Displays the full path of the generated `.intunewin` file upon success.
- Generates a Application in Intune
- The Install.bat file is used as information source for the Metadata for Intune. Use the REM lines
  Install.Bat       | Intune
  ------------------------------------
  REM DESCRIPTION   | Name
  REM MANUFACTURER  | Publischer
  REM LANGUAGE      | Not used
  REM FILENAME      | Executable used for detection rule
  REM VERSION       | App Version
  REM ASSETNUMBER   | Notes
  REM OWNER         | Owner of the App
- If you supply the executable file name for the installed Program and the Program is installed, a file detection rule is generated automatically.
- 

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
- If the output file already exists, a timestamp is appended to the filename to avoid overwriting.
- The script will display the full path of the generated file upon completion.

## Error Handling

- If the `IntuneWinAppUtil.exe` tool is missing, the script will automatically download it.
- Errors during execution are highlighted, and the script waits for user input before closing.

## Example Output

```plaintext
==========================================
          IntuneWin Packaging Tool         
==========================================

Packaging with IntuneWinAppUtil.exe...
File successfully packaged as:
C:\Path\To\Output\YourFolderName.intunewin

==========================================
Script completed successfully!
==========================================
```

## Notes

- Ensure that the `Install.bat` file exists in the source folder, as it is required by `IntuneWinAppUtil.exe`.
- The script automatically creates an `Output` directory in the same location as the script if it does not already exist.

## Troubleshooting

- **Missing Dependencies:** Ensure you have an active internet connection if the script needs to download `IntuneWinAppUtil.exe`.
- **File Not Found Errors:** Verify that the provided folder exists and contains the required files.

## License

This project is licensed under the MIT License. See the LICENSE file for details.

---

### Author

- **Your Name**  
- [GitHub Profile](https://github.com/YourUsername)

Feel free to contribute to this repository by submitting issues or pull requests!

