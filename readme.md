# File Manager in PowerShell

## Description

A graphical application for Windows written in PowerShell, designed for convenient viewing, sorting, and deleting files in a selected folder. Allows you to quickly work with large lists of files, supports deletion to Recycle Bin, opening files with a double-click, sorting by name, size, and date, as well as displaying statistics on the number and size of files.

---

## Installation

1. Download the `FileManager.ps1` file to any convenient folder.
2. Make sure you have PowerShell 5.1 (or newer) installed on your computer.
3. (Recommended) To run without a console window, create a shortcut with the `-WindowStyle Hidden` parameter or use a .vbs wrapper.

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

---

## Usage

- By default, the folder `G:\My Drive\recordings` is opened (if it exists).
- You can select another folder using the "Folder" button.
- The table displays files with extensions `.m4a`, `.mp3`, `.ogg`.
- Sorting by name, size, and date is available (buttons at the top of the window).
- To delete, select files and click "Delete" (you can choose to delete to Recycle Bin or permanently).
- Double-clicking a file opens it in the default application.
- At the bottom, statistics on the number and size of files are displayed.

---

## Configuration

- The default folder path can be changed in the `$global:folderPath` variable at the beginning of the script.
- Font size and family are set by the `$global:fontSize` and `$global:fontFamily` variables.
- To run without a console window, use a shortcut with the `-WindowStyle Hidden` parameter or a .vbs wrapper.

---

## Requirements

- Windows with PowerShell 5.1 or newer installed.
- Does not require third-party libraries.
- Does not make changes to the system.
- To run without a console window, it is recommended to use a shortcut or a .vbs wrapper.
