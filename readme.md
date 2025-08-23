# File Manager in PowerShell

A graphical application for Windows written in PowerShell, designed for convenient viewing, sorting, and deleting files in a selected folder. Supports audio file metadata editing, background processing, and enhanced file management.

**Note: This application requires TagLib installation to function.**

---

## 🚀 Features

- **File browsing and management** - View, sort, and delete files
- **Sorting options** - By name, duration, and creation date  
- **File operations** - Delete to Recycle Bin or permanently
- **Comments functionality** - Read and edit audio file metadata
- **Auto-loading metadata** - Automatic metadata loading with progress indicator
- **Background processing** - Non-blocking metadata loading for all files
- **Filtering options** - Filter by "Last 20" files or show "All" files
- **File types supported** - `.m4a`, `.mp3`, `.ogg` files
- **Display formats** - Duration in hh:mm:ss, dates as dd.MM.yy HH:mm:ss

---

## 🔧 Installation

### 1. PowerShell 7+ (Required)
```powershell
# Using winget (recommended)
winget install Microsoft.PowerShell
```

### 2. TagLibCli Module (Required)
```powershell
Install-Module -Name TagLibCli -Force
```

### 3. Verification
```powershell
$PSVersionTable.PSVersion  # Should show 7.x.x
Get-Module -Name TagLibCli -ListAvailable  # Should show the module
```

---

## 🖥️ Creating Desktop Shortcut

### Normal Mode (Console Visible)
1. **Right-click** on desktop → **New** → **Shortcut**
2. **Target path:**
   ```
   pwsh.exe -ExecutionPolicy Bypass -File "C:\path\to\FileManager.ps1"
   ```
   *(Replace `C:\path\to\FileManager.ps1` with actual file location)*
3. **Name:** File Manager

### Hidden Console Mode (Recommended)
1. **Right-click** on desktop → **New** → **Shortcut**
2. **Target path:**
   ```
   pwsh.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\path\to\FileManager.ps1"
   ```
   *(Replace `C:\path\to\FileManager.ps1` with actual file location)*
3. **Name:** File Manager
4. **Icon (optional):** Right-click shortcut → Properties → Change Icon → Browse to PowerShell icon

---

## 🔧 Troubleshooting

### TagLibCli Installation Issues
```powershell
# Check if module is installed
Get-Module -Name TagLibCli -ListAvailable

# Manual installation with force
Install-Module -Name TagLibCli -Force -AllowClobber

# Verify DLL exists (copy and paste as one line)
Test-Path (Join-Path (Split-Path (Get-Module -Name TagLibCli -ListAvailable).Path -Parent) "TagLibSharp.dll")
```
