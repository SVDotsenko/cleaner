# File Manager in PowerShell

A graphical application for Windows written in PowerShell, designed for convenient viewing, sorting, and deleting files in a selected folder. Allows you to quickly work with large lists of files, supports deletion to Recycle Bin, opening files with a double-click, sorting by name, size, and date, as well as displaying statistics on the number and size of files.

---

## 🚀 Quick Start - Basic Usage

**For users who want to quickly try the application without any setup**

### What You Get
- ✅ **File browsing and management** - View, sort, and delete files
- ✅ **Sorting options** - By name, size, and creation date  
- ✅ **File operations** - Delete to Recycle Bin or permanently
- ✅ **Statistics** - File count and total size display
- ✅ **No installation required** - Works with any PowerShell 5.1+

### Installation & Launch

1. **Download** the `FileManager.ps1` file to any folder
2. **Double-click** the file to run (if .ps1 files are associated with PowerShell)
3. **Or run via PowerShell:**
   ```powershell
   powershell.exe -ExecutionPolicy Bypass -File "C:\path\to\FileManager.ps1"
   ```

### Basic Usage

- **Default folder**: `G:\My Drive\recordings` (if it exists)
- **Select folder**: Use the "Folder" button to choose another directory
- **File types**: Displays `.m4a`, `.mp3`, `.ogg` files
- **Sorting**: Use the radio buttons at the top (Name, Size, Created)
- **Delete files**: Select files and click "Delete" 
- **Open files**: Double-click any file to open in default application
- **Display mode**: Compact mode with short file names and full date/time format

### Display Mode
- **Compact Mode**: Shortened file names with full date and time format (dd.MM.yy HH:mm:ss)

### Requirements
- Windows with PowerShell 5.1 or newer
- No additional software needed
- No system changes made

---

## 🔧 Advanced Features - Full Functionality

**For power users and developers who want the complete experience**

### What You Get (Everything Above +)
- ✅ **Comments functionality** - Read and edit audio file metadata
- ✅ **Auto-loading comments** - Automatic comment loading on startup and scroll
- ✅ **Real-time updates** - Comments displayed directly in file list
- ✅ **Scroll-based updates** - Comments automatically update when scrolling
- ✅ **Unit testing** - Comprehensive test suite for development
- ✅ **Enhanced UI feedback** - Loading indicators and better error handling

### Installation Requirements

#### 1. PowerShell 7+ (Required)
```powershell
# Option A - Using winget (recommended)
winget install Microsoft.PowerShell

# Option B - Manual download
# Go to https://github.com/PowerShell/PowerShell/releases
# Download latest version for Windows and install
```

#### 2. TagLibCli Module (Required)
```powershell
Install-Module -Name TagLibCli -Force
```

#### 3. Verification
```powershell
$PSVersionTable.PSVersion  # Should show 7.x.x
```
```powershell
Get-Module -Name TagLibCli -ListAvailable  # Should show the module
```

### Advanced Usage

#### Comments Feature
- **Automatic Loading**: Comments load automatically on application startup
- **Scroll-based Updates**: Comments automatically update when scrolling through the list
- **Manual Loading**: Use "Update" button to manually load comments for visible files
- **Individual Editing**: Select a single audio file to view/edit comments in text box
- **Real-time Display**: Comments shown directly in file list
- **Save Changes**: Edit comments and click "Save" to update files
- **Supported Formats**: MP3, M4A, OGG audio files

#### Enhanced Sorting
- **Name**: Sorts by file name (A-Z), then by date (newest first)
- **Size**: Sorts by file size (largest first)  
- **Created**: Sorts by creation date (newest first)

#### Scroll Detection
- **Mouse wheel scrolling**: Automatically updates comments
- **Keyboard navigation**: Arrow keys, Page Up/Down, Home/End trigger updates
- **Scrollbar dragging**: Mouse interactions with scrollbar trigger updates
- **Debounced updates**: 500ms delay after scrolling stops to prevent excessive updates

### Testing (For Developers)

#### Prerequisites
- **PowerShell 7+** (required)
- **Pester 5.x** (required)

#### Installing Pester 5.x
```powershell
# Check current version
Get-Module -Name Pester -ListAvailable
```
```powershell
# Remove old versions (if installed via PowerShellGet)
Uninstall-Module -Name Pester -AllVersions -Force
```
```powershell
# For system-installed Pester, manually remove (requires admin):
Remove-Item -Path "C:\Program Files\WindowsPowerShell\Modules\Pester" -Recurse -Force
```
```powershell
# Install latest Pester 5.x
Install-Module -Name Pester -Force -SkipPublisherCheck
```

**Alternative approach (without removing system module):**
```powershell
# Install Pester 5.x alongside existing version
Install-Module -Name Pester -Force -SkipPublisherCheck
```
```powershell
# Force load newer version when running tests
Import-Module Pester -Force
```

#### Running Tests
**Important**: Run from project root directory (where `FileManager.ps1` is located)

```powershell
# Navigate to project root
cd C:\repositories\cleaner
```
```powershell
# Run all tests
Invoke-Pester -Path ".\tests\"
```
```powershell
# Run with detailed output
Invoke-Pester -Path ".\tests\" -Output Detailed
```
```powershell
# Force reload Pester module (if having issues)
Import-Module Pester -Force; Invoke-Pester -Path ".\tests\"
```

#### Test Coverage
Tests are located in `tests/FileManager.Tests.ps1` and cover:
- **Format-ExtractedDate**: Date formatting with/without time, null handling
- **Get-DisplayNameFromFileName**: Filename processing and letter extraction

#### Troubleshooting Tests
If tests fail with Pester 3.x errors:
```powershell
# Check version
Get-Module -Name Pester
```
```powershell
# If showing 3.x, force reload 5.x
Import-Module Pester -Force
```
```powershell
Invoke-Pester -Path ".\tests\"
```

### Configuration

- **Default folder**: Change `$global:folderPath` variable at script beginning
- **Font settings**: Modify `$global:fontSize` and `$global:fontFamily` variables
- **Console window**: Use shortcut with `-WindowStyle Hidden` parameter

### Requirements for Advanced Features
- **PowerShell 7+** (required for comments)
- **TagLibCli module** (required for comments)
- **Audio files** with supported formats (MP3, M4A, OGG)
- **Pester 5.x** (required for testing)

If requirements are not met, the application will show notification messages and disable the comments interface.