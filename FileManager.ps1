Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

$global:fontSize = 14
$global:fontFamily = "Segoe UI"

$form = New-Object Windows.Forms.Form
$form.Text = "File Manager"
$form.Width = 1200
$form.Height = 800
$form.MinimumSize = New-Object Drawing.Size(600,400)
$form.StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen

$controls = @{}

$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 5000
$toolTip.InitialDelay = 1000
$toolTip.ReshowDelay = 500

function CreateControls {
    $gap = [int]($global:fontSize * 0.8)
    $btnH = [int]($global:fontSize * 2.2)
    $y = $gap
    $x = $gap
    $controls.SelectFolder = New-Object Windows.Forms.Button
    $controls.SelectFolder.Text = "Folder"
    $form.Controls.Add($controls.SelectFolder)
    $controls.SelectFolder.SetBounds($x, $y, 100 + $global:fontSize*2, $btnH)
    $x += $controls.SelectFolder.Width + $gap
    $controls.DeleteBtn = New-Object Windows.Forms.Button
    $controls.DeleteBtn.Text = "Delete"
    $controls.DeleteBtn.Enabled = $false
    $form.Controls.Add($controls.DeleteBtn)
    $controls.DeleteBtn.SetBounds($x, $y, 120 + $global:fontSize*2, $btnH)
    $x += $controls.DeleteBtn.Width + $gap
    $controls.DeleteToTrashCheckBox = New-Object Windows.Forms.CheckBox
    $controls.DeleteToTrashCheckBox.Checked = $true
    $controls.DeleteToTrashCheckBox.AutoSize = $true
    $form.Controls.Add($controls.DeleteToTrashCheckBox)
    $controls.DeleteToTrashCheckBox.SetBounds($x, $y, 100 + $global:fontSize*2, $btnH)
    $x += $controls.DeleteToTrashCheckBox.Width + $gap
    $controls.SortNameBtn = New-Object Windows.Forms.Button
    $controls.SortNameBtn.Text = "Name"
    $form.Controls.Add($controls.SortNameBtn)
    $controls.SortNameBtn.SetBounds($x, $y, 110 + $global:fontSize*2, $btnH)
    $x += $controls.SortNameBtn.Width + $gap
    $controls.SortSizeBtn = New-Object Windows.Forms.Button
    $controls.SortSizeBtn.Text = "Size"
    $form.Controls.Add($controls.SortSizeBtn)
    $controls.SortSizeBtn.SetBounds($x, $y, 110 + $global:fontSize*2, $btnH)
    $x += $controls.SortSizeBtn.Width + $gap
    $controls.SortCreatedBtn = New-Object Windows.Forms.Button
    $controls.SortCreatedBtn.Text = "Created"
    $form.Controls.Add($controls.SortCreatedBtn)
    $controls.SortCreatedBtn.SetBounds($x, $y, 110 + $global:fontSize*2, $btnH)
    $x += $controls.SortCreatedBtn.Width + $gap
    $y2 = $y + $btnH + $gap
    
    # Create StatusStrip instead of labels
    $controls.StatusStrip = New-Object Windows.Forms.StatusStrip
    $controls.StatusLabel = New-Object Windows.Forms.ToolStripStatusLabel
    $controls.StatusLabel.Text = "Total files: 0 | Total size: 0 MB"
    $controls.StatusStrip.Items.Add($controls.StatusLabel)
    $form.Controls.Add($controls.StatusStrip)
    
    $controls.ListView = New-Object Windows.Forms.ListView
    $controls.ListView.View = 'Details'
    $controls.ListView.FullRowSelect = $true
    $controls.ListView.MultiSelect = $true
    $controls.ListView.Scrollable = $true
    $controls.ListView.GridLines = $true
    $controls.ListView.Sorting = 'None'
    $controls.ListView.Anchor = "Top,Bottom,Left,Right"
    $form.Controls.Add($controls.ListView)
    $controls.ListView.Left = $gap
    $controls.ListView.Top = $y2 + $btnH + $gap
    $controls.ListView.Width = $form.ClientSize.Width - $gap*2
    $controls.ListView.Height = $form.ClientSize.Height - $controls.ListView.Top - $gap
    $controls.ListView.Columns.Add("File Name", -1) | Out-Null
    $controls.ListView.Columns.Add("MB", 100) | Out-Null
    $controls.ListView.Columns.Add("Created", 100) | Out-Null
    $selectFolderTooltip = @"
Allows you to select another folder to display and work with its files.
"@
    $deleteTooltip = @"
Deletes the selected files from the table.
"@
    $deleteToTrashTooltip = @"
When checked, files are moved to the Recycle Bin. When unchecked, files are permanently deleted.
"@
    $sortNameTooltip = @"
Sorts the file list by name (alphabetically, A to Z).
"@
    $sortSizeTooltip = @"
Sorts the file list by size (from largest to smallest).
"@
    $sortCreatedTooltip = @"
Sorts the file list by creation date (newest first).
"@
    $toolTip.SetToolTip($controls.SelectFolder, $selectFolderTooltip.Trim())
    $toolTip.SetToolTip($controls.DeleteBtn, $deleteTooltip.Trim())
    $toolTip.SetToolTip($controls.DeleteToTrashCheckBox, $deleteToTrashTooltip.Trim())
    $toolTip.SetToolTip($controls.SortNameBtn, $sortNameTooltip.Trim())
    $toolTip.SetToolTip($controls.SortSizeBtn, $sortSizeTooltip.Trim())
    $toolTip.SetToolTip($controls.SortCreatedBtn, $sortCreatedTooltip.Trim())
    $controls.ShowFullNameCheckBox = New-Object Windows.Forms.CheckBox
    $controls.ShowFullNameCheckBox.Text = "Show full name"
    $controls.ShowFullNameCheckBox.Checked = $true
    $controls.ShowFullNameCheckBox.AutoSize = $true
    $form.Controls.Add($controls.ShowFullNameCheckBox)
    $controls.ShowFullNameCheckBox.SetBounds($x, $y, 140 + $global:fontSize*2, $btnH)
    $x += $controls.ShowFullNameCheckBox.Width + $gap
}

function LayoutOnlyFonts {
    $gap = [int]($global:fontSize * 0.8)
    $btnH = [int]($global:fontSize * 2.2)
    $y = $gap
    $x = $gap
    $controls.SelectFolder.SetBounds($x, $y, 100 + $global:fontSize*2, $btnH)
    $x += $controls.SelectFolder.Width + $gap
    $controls.DeleteBtn.SetBounds($x, $y, 120 + $global:fontSize*2, $btnH)
    $x += $controls.DeleteBtn.Width + $gap
    $controls.DeleteToTrashCheckBox.SetBounds($x, $y, 100 + $global:fontSize*2, $btnH)
    $x += $controls.DeleteToTrashCheckBox.Width + $gap
    $controls.SortNameBtn.SetBounds($x, $y, 110 + $global:fontSize*2, $btnH)
    $x += $controls.SortNameBtn.Width + $gap
    $controls.SortSizeBtn.SetBounds($x, $y, 110 + $global:fontSize*2, $btnH)
    $x += $controls.SortSizeBtn.Width + $gap
    $controls.SortCreatedBtn.SetBounds($x, $y, 110 + $global:fontSize*2, $btnH)
    $x += $controls.SortCreatedBtn.Width + $gap
    $y2 = $y + $btnH + $gap
    $controls.StatusLabel.Font = $font
    $gap = [int]($global:fontSize * 0.8)
    $controls.ListView.Left = $gap
    $controls.ListView.Top = $y2 + $btnH + $gap
    $controls.ListView.Width = $form.ClientSize.Width - $gap*2
    $controls.ListView.Height = $form.ClientSize.Height - $controls.ListView.Top - $gap
    $form.PerformLayout()
    $controls.ListView.Font = $font
    $controls.SelectFolder.Font = $font
    $controls.DeleteBtn.Font = $font
    $controls.DeleteToTrashCheckBox.Font = $font
    $controls.SortNameBtn.Font = $font
    $controls.SortSizeBtn.Font = $font
    $controls.SortCreatedBtn.Font = $font
    $controls.StatusLabel.Font = $font
}

function Set-AllFonts($fontSize) {
    $font = New-Object System.Drawing.Font($global:fontFamily, $fontSize)
    foreach ($ctrl in $controls.Values) { $ctrl.Font = $font }
    $form.Font = $font
}

function Update-InfoLabels {
    $selectedCount = $controls.ListView.SelectedItems.Count
    if ($selectedCount -gt 0) {
        $sum = 0
        foreach ($selectedItem in $controls.ListView.SelectedItems) {
            $fileIndex = $selectedItem.Index
            $file = $global:filteredTable[$fileIndex]
            $sum += $file.SizeMB
        }
        $sum = [math]::Round($sum, 2)
        $controls.StatusLabel.Text = "Selected files: $selectedCount | Selected size: $sum MB"
    } else {
        $count = $global:filteredTable.Count
        $sum = 0
        if ($count -gt 0) {
            $sum = ($global:filteredTable | Measure-Object -Property SizeMB -Sum).Sum
            $sum = [math]::Round($sum, 2)
        }
        $controls.StatusLabel.Text = "Total files: $count | Total size: $sum MB"
    }
}

function Update-ListView {
    $controls.ListView.Items.Clear()
    foreach ($file in $global:filteredTable) {
        # Show names and dates based on current checkbox state
        $displayName = if ($controls.ShowFullNameCheckBox.Checked) { $file.OrigName } else { $file.Name }
        $item = New-Object Windows.Forms.ListViewItem($displayName)
        $item.SubItems.Add("$($file.SizeMB)")
        $displayDate = Format-ExtractedDate $file.DisplayDate $controls.ShowFullNameCheckBox.Checked
        $item.SubItems.Add($displayDate)
        $controls.ListView.Items.Add($item) | Out-Null
    }
    $controls.DeleteBtn.Enabled = $controls.ListView.Items.Count -gt 0 -and $controls.ListView.SelectedItems.Count -gt 0
    Update-InfoLabels
    Update-ListViewTextColors
}

function Update-ListViewPreserveScroll {
    # Save current scroll position and selected items
    $topItemIndex = -1
    if ($null -ne $controls.ListView.TopItem) {
        $topItemIndex = $controls.ListView.TopItem.Index
    }
    
    $selectedIndexes = @()
    foreach ($item in $controls.ListView.SelectedItems) {
        $selectedIndexes += $item.Index
    }
    
    # Suspend layout updates to prevent flickering
    $controls.ListView.BeginUpdate()
    
    $controls.ListView.Items.Clear()
    foreach ($file in $global:filteredTable) {
        $displayName = if ($controls.ShowFullNameCheckBox.Checked) { $file.OrigName } else { $file.Name }
        $item = New-Object Windows.Forms.ListViewItem($displayName)
        $item.SubItems.Add("$($file.SizeMB)")
        $displayDate = Format-ExtractedDate $file.DisplayDate $controls.ShowFullNameCheckBox.Checked
        $item.SubItems.Add($displayDate)
        $controls.ListView.Items.Add($item) | Out-Null
    }
    
    # Restore scroll position if possible
    if ($topItemIndex -ge 0 -and $topItemIndex -lt $controls.ListView.Items.Count) {
        $controls.ListView.TopItem = $controls.ListView.Items[$topItemIndex]
    }
    
    # Restore selected items if possible
    foreach ($index in $selectedIndexes) {
        if ($index -lt $controls.ListView.Items.Count) {
            $controls.ListView.Items[$index].Selected = $true
        }
    }
    
    # Resume layout updates
    $controls.ListView.EndUpdate()
    
    $controls.DeleteBtn.Enabled = $controls.ListView.Items.Count -gt 0 -and $controls.ListView.SelectedItems.Count -gt 0
    Update-InfoLabels
    Update-ListViewTextColors
}

function Update-ListViewTextColors {
    # Update text display based on ShowFullName checkbox state
    foreach ($item in $controls.ListView.Items) {
        $fileIndex = $item.Index
        if ($fileIndex -lt $global:filteredTable.Count) {
            $file = $global:filteredTable[$fileIndex]
            
            # Update name column text
            if ($controls.ShowFullNameCheckBox.Checked) {
                # Show full name
                $item.Text = $file.OrigName
                $item.ForeColor = [System.Drawing.Color]::Black
            } else {
                # Show short name
                $item.Text = $file.Name
                $item.ForeColor = [System.Drawing.Color]::Black
            }
            
            # Update date column text
            if ($controls.ShowFullNameCheckBox.Checked) {
                # Show full date with time
                $item.SubItems[2].Text = Format-ExtractedDate $file.DisplayDate $true
                $item.SubItems[2].ForeColor = [System.Drawing.Color]::Black
            } else {
                # Show short date without time
                $item.SubItems[2].Text = Format-ExtractedDate $file.DisplayDate $false
                $item.SubItems[2].ForeColor = [System.Drawing.Color]::Black
            }
            
            # Size column always black
            $item.SubItems[1].ForeColor = [System.Drawing.Color]::Black
        }
    }
    
    # Auto-resize columns and remember maximum widths
    $controls.ListView.AutoResizeColumn(0, [System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
    $controls.ListView.AutoResizeColumn(2, [System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
    
    # Remember maximum widths
    $currentNameWidth = $controls.ListView.Columns[0].Width
    $currentDateWidth = $controls.ListView.Columns[2].Width
    
    if ($currentNameWidth -gt $global:maxColumnWidths.Name) {
        $global:maxColumnWidths.Name = $currentNameWidth
    }
    if ($currentDateWidth -gt $global:maxColumnWidths.Date) {
        $global:maxColumnWidths.Date = $currentDateWidth
    }
    
    # Set columns to maximum width if we have stored values
    if ($global:maxColumnWidths.Name -gt 0) {
        $controls.ListView.Columns[0].Width = $global:maxColumnWidths.Name
    }
    if ($global:maxColumnWidths.Date -gt 0) {
        $controls.ListView.Columns[2].Width = $global:maxColumnWidths.Date
    }
}

function Get-DateFromFileName($fileName, $extension) {
    $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    switch ($extension.ToLower()) {
        ".m4a" {
            if ($nameWithoutExt -match ".*?(\d{6})_(\d{6})$") {
                $dateStr = $matches[1]
                $timeStr = $matches[2]
                $year = "20" + $dateStr.Substring(0, 2)
                $month = $dateStr.Substring(2, 2)
                $day = $dateStr.Substring(4, 2)
                $hour = $timeStr.Substring(0, 2)
                $minute = $timeStr.Substring(2, 2)
                $second = $timeStr.Substring(4, 2)
                try {
                    return [DateTime]::ParseExact("$year-$month-$day $hour`:$minute`:$second", 'yyyy-MM-dd HH:mm:ss', $null)
                } catch {
                    return $null
                }
            }
        }
        ".ogg" {
            if ($nameWithoutExt -match ".*?(\d{4}_\d{2}_\d{2})_(\d{2}_\d{2}_\d{2})") {
                $dateStr = $matches[1]
                $timeStr = $matches[2]
                try {
                    $dateTime = [DateTime]::ParseExact("$dateStr $timeStr", 'yyyy_MM_dd HH_mm_ss', $null)
                    return $dateTime
                } catch {
                    return $null
                }
            }
        }
        ".mp3" {
            if ($nameWithoutExt -match "(20\d{6})(\d{6})") {
                $dateStr = $matches[1]
                $timeStr = $matches[2]
                $year = $dateStr.Substring(0, 4)
                $month = $dateStr.Substring(4, 2)
                $day = $dateStr.Substring(6, 2)
                $hour = $timeStr.Substring(0, 2)
                $minute = $timeStr.Substring(2, 2)
                $second = $timeStr.Substring(4, 2)
                try {
                    return [DateTime]::ParseExact("$year-$month-$day $hour`:$minute`:$second", 'yyyy-MM-dd HH:mm:ss', $null)
                } catch {
                    return $null
                }
            }
        }
    }
    return $null
}

function Format-ExtractedDate($date, $showTime = $false) {
    if ($null -eq $date) {
        return "N/A"
    }
    if ($showTime) {
        return $date.ToString("dd.MM.yy HH:mm:ss")
    } else {
        return $date.ToString("dd.MM.yy")
    }
}

function Get-DisplayNameFromFileName($fileName) {
    $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    if ($nameWithoutExt.Length -gt 0 -and [char]::IsLetter($nameWithoutExt[0])) {
        $result = ""
        foreach ($c in $nameWithoutExt.ToCharArray()) {
            if ([char]::IsLetter($c)) {
                $result += $c
            } else {
                break
            }
        }
        return $result
    } else {
        return $fileName
    }
}

$global:folderPath = "G:\My Drive\recordings"
$global:fileTable = @()
$global:filteredTable = @()
$global:activeSortButton = $null
$global:maxColumnWidths = @{ Name = 0; Date = 0 }

function Show-TrayNotification {
    param(
        [string]$Title,
        [string]$Message,
        [int]$Duration = 3000,
        [string]$Type = "Info"
    )
    
    # Log to console
    Write-Host "[$Type] $Title`: $Message"
    
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $notifyIcon = New-Object System.Windows.Forms.NotifyIcon
        
        # Set icon based on type
        switch ($Type) {
            "Error" { 
                $notifyIcon.Icon = [System.Drawing.SystemIcons]::Error
                $toolTipIcon = [System.Windows.Forms.ToolTipIcon]::Error
            }
            default { 
                $notifyIcon.Icon = [System.Drawing.SystemIcons]::Information
                $toolTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
            }
        }
        
        $notifyIcon.Visible = $true
        $notifyIcon.ShowBalloonTip($Duration, $Title, $Message, $toolTipIcon)
        Start-Sleep -Milliseconds $Duration
        $notifyIcon.Dispose()
    } catch {
        $messageBoxIcon = if ($Type -eq "Error") { 'Error' } else { 'Information' }
        [System.Windows.Forms.MessageBox]::Show($Message, $Title, 'OK', $messageBoxIcon) | Out-Null
    }
}

function Rename-CallRecordingFiles {
    if (-not (Test-Path $global:folderPath)) {
        return
    }
    $files = Get-ChildItem -Path $global:folderPath -File
    $renamedCount = 0
    foreach ($file in $files) {
        if ($file.Name.StartsWith("Call recording ")) {
            $newName = $file.Name.Substring(15)
            if (-not [string]::IsNullOrWhiteSpace($newName)) {
                $newPath = Join-Path $file.Directory.FullName $newName
                try {
                    if (-not (Test-Path $newPath)) {
                        Rename-Item -Path $file.FullName -NewName $newName -Force
                        $renamedCount++
                    }
                } catch {
                }
            }
        }
    }
    if ($renamedCount -gt 0) {
        Show-TrayNotification -Title "File Manager" -Message "$renamedCount file(s) renamed (removed 'Call recording ' prefix)."
    }
}

function Get-FilesFromFolder {
    Rename-CallRecordingFiles
    $global:fileTable = @()
    if (Test-Path $global:folderPath) {
        $files = Get-ChildItem -Path $global:folderPath -File | Where-Object { $_.Extension -match "\.(m4a|mp3|ogg)$" }
        foreach ($file in $files) {
            $extractedDate = Get-DateFromFileName $file.Name $file.Extension
            $displayDate = if ($null -eq $extractedDate) { $file.CreationTime } else { $extractedDate }
            $displayName = Get-DisplayNameFromFileName $file.Name
            $global:fileTable += [PSCustomObject]@{
                Name   = $displayName
                SizeMB = [math]::Round($file.Length / 1MB, 2)
                Path   = $file.FullName
                CreationTime = $file.CreationTime
                ExtractedDate = $extractedDate
                DisplayDate = $displayDate
                OrigName = $file.Name
            }
        }
        $global:activeSortButton = $controls.SortCreatedBtn
        $global:fileTable = $global:fileTable | Sort-Object DisplayDate -Descending
        Update-SortButtonStates
        Invoke-Search
    } else {
        $global:fileTable = @()
        $global:activeSortButton = $controls.SortCreatedBtn
        Update-SortButtonStates
        Invoke-Search
    }
}

function Invoke-Search {
    $pattern = $controls.SearchBox.Text
    if ([string]::IsNullOrWhiteSpace($pattern)) {
        $global:filteredTable = $global:fileTable
    } else {
        try {
            $regex = New-Object System.Text.RegularExpressions.Regex($pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            $global:filteredTable = $global:fileTable | Where-Object { $regex.IsMatch([System.IO.Path]::GetFileNameWithoutExtension($_.Name)) }
        } catch {
            $p = $pattern.ToLower()
            $global:filteredTable = $global:fileTable | Where-Object { $_.Name.ToLower() -like "*$p*" }
        }
    }
    Update-ListView
}

function Move-FileToRecycleBin($filePath) {
    try {
        $shell = New-Object -ComObject Shell.Application
        $item = $shell.Namespace(0).ParseName($filePath)
        $item.InvokeVerb("delete")
        return $true
    } catch {
        return $false
    }
}

function Update-SortButtonStates {
    # Сначала сбрасываем состояние всех кнопок
    $controls.SortNameBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
    $controls.SortNameBtn.BackColor = [System.Drawing.SystemColors]::Control
    $controls.SortNameBtn.Enabled = $true
    $controls.SortSizeBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
    $controls.SortSizeBtn.BackColor = [System.Drawing.SystemColors]::Control
    $controls.SortSizeBtn.Enabled = $true
    $controls.SortCreatedBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
    $controls.SortCreatedBtn.BackColor = [System.Drawing.SystemColors]::Control
    $controls.SortCreatedBtn.Enabled = $true

    # Затем выделяем активную кнопку и отключаем её
    if ($null -ne $global:activeSortButton) {
        $global:activeSortButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $global:activeSortButton.BackColor = [System.Drawing.Color]::LightBlue
        $global:activeSortButton.ForeColor = [System.Drawing.Color]::DarkBlue
        $global:activeSortButton.Enabled = $false
    }
}

function BindHandlers {
    $controls.SelectFolder.Add_Click({
        $dialog = New-Object Windows.Forms.FolderBrowserDialog
        $dialog.Description = "Select a folder"
        if ($dialog.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) {
            $global:folderPath = $dialog.SelectedPath
            Get-FilesFromFolder
        }
    })
    $controls.ListView.Add_SelectedIndexChanged({ 
        $controls.DeleteBtn.Enabled = $controls.ListView.SelectedItems.Count -gt 0
        Update-InfoLabels
    })
    $controls.SortNameBtn.Add_Click({ 
        $global:activeSortButton = $controls.SortNameBtn
        Update-SortButtonStates
        $global:filteredTable = $global:filteredTable | Sort-Object @{Expression="Name"; Ascending=$true}, @{Expression="DisplayDate"; Ascending=$false}
        Update-ListView 
    })
    $controls.SortSizeBtn.Add_Click({ 
        $global:activeSortButton = $controls.SortSizeBtn
        Update-SortButtonStates
        $global:filteredTable = $global:filteredTable | Sort-Object SizeMB -Descending
        Update-ListView 
    })
    $controls.SortCreatedBtn.Add_Click({ 
        $global:activeSortButton = $controls.SortCreatedBtn
        Update-SortButtonStates
        $global:filteredTable = $global:filteredTable | Sort-Object DisplayDate -Descending
        Update-ListView 
    })
    $controls.DeleteBtn.Add_Click({
        $toDeleteIndexes = @()
        foreach ($item in $controls.ListView.SelectedItems) {
            $toDeleteIndexes += $item.Index
        }
        $toDeleteIndexes = $toDeleteIndexes | Sort-Object -Descending
        $deleted = 0
        $useTrash = $controls.DeleteToTrashCheckBox.Checked
        foreach ($i in $toDeleteIndexes) {
            $file = $global:filteredTable[$i]
            $success = $false
            try {
                if ($useTrash) {
                    $success = Move-FileToRecycleBin $file.Path
                } else {
                    Remove-Item -Path $file.Path -Force
                    $success = $true
                }
                if ($success) {
                    $deleted++
                    $global:fileTable = $global:fileTable | Where-Object { $_.Path -ne $file.Path }
                    $global:filteredTable = $global:filteredTable | Where-Object { $_.Path -ne $file.Path }
                }
            } catch {
            }
        }
        Update-ListViewPreserveScroll
        $actionText = if ($useTrash) { "moved to Recycle Bin" } else { "permanently deleted" }
        Show-TrayNotification -Title "Done" -Message "$deleted file(s) $actionText."
    })
    $controls.ListView.Add_DoubleClick({
        if ($controls.ListView.SelectedItems.Count -eq 1) {
            $index = $controls.ListView.SelectedItems[0].Index
            $file = $global:filteredTable[$index]
            try {
                [System.Diagnostics.Process]::Start($file.Path) | Out-Null
            } catch {
                Show-TrayNotification -Title "Error" -Message "Cannot open file: $($file.Path)" -Type "Error"
            }
        }
    })
    $controls.ShowFullNameCheckBox.Add_CheckedChanged({ Update-ListViewTextColors })
    $controls.DeleteToTrashCheckBox.Add_CheckedChanged({
        if ($controls.DeleteToTrashCheckBox.Checked) {
            $controls.DeleteBtn.Text = "Bin"
        } else {
            $controls.DeleteBtn.Text = "Delete"
        }
    })
}

$form.Add_Resize({
    if ($null -ne $controls.ListView) {
        $gap = [int]($global:fontSize * 0.8)
        $controls.ListView.Left = $gap
        $controls.ListView.Width = $form.ClientSize.Width - $gap*2
        $controls.ListView.Height = $form.ClientSize.Height - $controls.ListView.Top - $gap
        $controls.ListView.AutoResizeColumn(0, [System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
    }
})

$form.Topmost = $false

$form.Add_Shown({
    CreateControls
    Set-AllFonts $global:fontSize
    BindHandlers
    # Установим начальный текст кнопки удаления в соответствии с чекбоксом
    if ($controls.DeleteToTrashCheckBox.Checked) {
        $controls.DeleteBtn.Text = "Bin"
    } else {
        $controls.DeleteBtn.Text = "Delete"
    }
    Get-FilesFromFolder
    $form.Activate()
})

[void]$form.ShowDialog()
