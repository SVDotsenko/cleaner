# File Manager in PowerShell

## Description

A graphical application for Windows written in PowerShell, designed for convenient viewing, sorting, and deleting files in a selected folder. Allows you to quickly work with large lists of files, supports deletion to Recycle Bin, opening files with a double-click, sorting by name, size, and date, as well as displaying statistics on the number and size of files.

**New Feature**: Comments functionality allows you to read and edit comments/metadata in audio files (MP3, M4A, OGG).

---

## Installation

1. Download the `FileManager.ps1` file to any convenient folder.
2. Make sure you have PowerShell 5.1 (or newer) installed on your computer.
3. (Recommended) To run without a console window, create a shortcut with the `-WindowStyle Hidden` parameter or use a .vbs wrapper.

### Additional Requirements for Comments Functionality

To use the comments feature (reading and editing file comments), you need to install additional components:

#### 1. PowerShell 7+ (Required)
The comments functionality requires PowerShell 7 or newer. To install:

**Option A - Using winget (recommended):**
```
winget install Microsoft.PowerShell
```

**Option B - Manual download:**
- Go to https://github.com/PowerShell/PowerShell/releases
- Download the latest version for Windows
- Install and restart your computer

#### 2. TagLibCli Module (Required)
Install the TagLibCli PowerShell module:

```powershell
Install-Module -Name TagLibCli -Force
```

**Note**: This module includes the necessary TagLibSharp.dll file for reading/writing audio file metadata.

#### 3. Verification
After installation, you can verify everything is working by running:
```powershell
$PSVersionTable.PSVersion  # Should show 7.x.x
Get-Module -Name TagLibCli -ListAvailable  # Should show the module
```

---

## Launch

- Double-click the `FileManager.ps1` file (if .ps1 files are associated with PowerShell).
- Or run via PowerShell with the command:
  ```
  powershell.exe -ExecutionPolicy Bypass -File "C:\path\to\FileManager.ps1"
  ```
- To run without a console window, use a shortcut:
  ```
  powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\path\to\FileManager.ps1"
  ```

**Important**: For comments functionality, make sure you're running with PowerShell 7+.

---

## Usage

- By default, the folder `G:\My Drive\recordings` is opened (if it exists).
- You can select another folder using the "Folder" button.
- The table displays files with extensions `.m4a`, `.mp3`, `.ogg`.
- Sorting by name, size, and date is available (buttons at the top of the window).
- To delete, select files and click "Delete" (you can choose to delete to Recycle Bin or permanently).
- Double-clicking a file opens it in the default application.
- At the bottom, statistics on the number and size of files are displayed.

### Comments Feature
- Select a single audio file to view its comments in the text box below the "Folder" button.
- Edit the comments and click "Update" to save changes.
- The feature only works with audio files (MP3, M4A, OGG).
- If requirements are not met, you'll see notification messages and the comments interface will be hidden.

---

## Configuration

- The default folder path can be changed in the `$global:folderPath` variable at the beginning of the script.
- Font size and family are set by the `$global:fontSize` and `$global:fontFamily` variables.
- To run without a console window, use a shortcut with the `-WindowStyle Hidden` parameter or a .vbs wrapper.

---

## Requirements

- Windows with PowerShell 5.1 or newer installed.
- Does not require third-party libraries for basic functionality.
- Does not make changes to the system.
- To run without a console window, it is recommended to use a shortcut or a .vbs wrapper.

### For Comments Functionality:
- **PowerShell 7+** (required)
- **TagLibCli module** (required)
- Audio files with supported formats (MP3, M4A, OGG)

If requirements are not met, the application will show notification messages and disable the comments interface.
