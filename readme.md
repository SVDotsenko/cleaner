# File Manager in PowerShell

## Description

A graphical application for Windows written in PowerShell, designed for convenient viewing, sorting, and deleting files in a selected folder. Allows you to quickly work with large lists of files, supports deletion to Recycle Bin, opening files with a double-click, sorting by name, size, and date, as well as displaying statistics on the number and size of files.

**New Feature**: Comments functionality allows you to read and edit comments/metadata in audio files (MP3, M4A, OGG). The application now supports automatic loading of comments when switching to short name mode and displays comments directly in the file list.

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

## Testing

The project includes unit tests written for **Pester 5.x** and **PowerShell 7+**.

### Checking Pester Version

To check which version of Pester is installed:

```powershell
# Check all installed Pester versions
Get-Module -Name Pester -ListAvailable

# Check currently loaded Pester version
Get-Module -Name Pester

# Check only version number
(Get-Module -Name Pester).Version
```

### Installing/Updating Pester

If you have an older version of Pester (3.x), update to version 5.x:

```powershell
# Remove old versions (if installed via PowerShellGet)
Uninstall-Module -Name Pester -AllVersions -Force

# For system-installed Pester (like 3.4.0), manually remove it:
# Navigate to the module directory and delete the Pester folder
# Note: This requires administrator privileges
Remove-Item -Path "C:\Program Files\WindowsPowerShell\Modules\Pester" -Recurse -Force

# Install latest Pester 5.x
Install-Module -Name Pester -Force -SkipPublisherCheck

# Verify installation
Get-Module -Name Pester -ListAvailable
```

**Alternative approach (without removing system module):**
```powershell
# Install Pester 5.x alongside the existing version
Install-Module -Name Pester -Force -SkipPublisherCheck

# Force load the newer version when running tests
Import-Module Pester -Force
Invoke-Pester -Path ".\tests\"
```

**Note**: If you get "Access Denied" when trying to remove the system Pester module, you need to run PowerShell as Administrator. Right-click on PowerShell and select "Run as Administrator", then execute the removal command.
```

### Running Tests

**Important**: Tests must be run from the **project root directory** (where `FileManager.ps1` is located), not from the `tests` folder.

```powershell
# Navigate to project root
cd C:\repositories\cleaner

# Run all tests
Invoke-Pester -Path ".\tests\"

# Run with detailed output
Invoke-Pester -Path ".\tests\" -Output Detailed

# Force reload Pester module (if having issues)
Import-Module Pester -Force; Invoke-Pester -Path ".\tests\"
```

### Test Structure

Tests are located in `tests/FileManager.Tests.ps1` and cover:

- **Format-ExtractedDate**: Date formatting with/without time, null handling
- **Get-DisplayNameFromFileName**: Filename processing and letter extraction

### Troubleshooting

If tests fail with Pester 3.x errors, ensure you're using Pester 5.x:

```powershell
# Check version
Get-Module -Name Pester

# If showing 3.x, force reload 5.x
Import-Module Pester -Force
Invoke-Pester -Path ".\tests\"
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

### Display Modes
- **Full Name Mode**: Shows complete file names and full dates with time
- **Short Name Mode**: Shows shortened file names and dates without time, plus a Comments column
- Toggle between modes using the "Short name"/"Full name" button

### Comments Feature
- **Automatic Loading**: When switching to short name mode, comments are automatically loaded for visible files
- **Manual Loading**: Use the "Update" button to load comments for visible files
- **Individual Selection**: Select a single audio file to view and edit its comments in the text box
- **Real-time Updates**: Comments are displayed directly in the file list when in short name mode
- **Save Changes**: Edit comments in the text box and click "Save" to update the file
- **Supported Formats**: MP3, M4A, OGG audio files
- **Requirements Check**: The application automatically checks for required components and shows notifications

### Sorting Options
- **Name**: Sorts by file name (A-Z), then by date (newest first)
- **Size**: Sorts by file size (largest first)
- **Created**: Sorts by creation date (newest first)

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

### For Testing:
- **PowerShell 7+** (required)
- **Pester 5.x** (required)

---

## Recent Updates

- **Auto-loading comments** when switching to short name mode
- **Comments column** in the file list for quick viewing
- **Improved sorting** by name (now sorts by name then by date)
- **Better UI feedback** during comment loading operations
- **Enhanced error handling** and user notifications
- **Real-time comment updates** in the file list
- **Unit tests** for core functions using Pester 5.x
