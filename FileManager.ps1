Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

$global:folderPath = "G:\My Drive\recordings"
$global:fontSize = 14
$global:fontFamily = "Segoe UI"
$global:commentsEnabled = $false
$global:fileTable = @()
$global:filteredTable = @()
$global:activeSortButton = $null
$global:maxColumnWidths = @{ Name = 0; Date = 0; Comments = 0 }
$global:currentSelectedFile = $null
$global:originalCommentsText = ""
$global:lastScrollTop = 0
$global:scrollTimer = $null

$form = New-Object Windows.Forms.Form
$form.Text = "File Manager"
$form.Width = 1400
$form.Height = 800
$form.MinimumSize = New-Object Drawing.Size(800,400)
$form.StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen

$controls = @{}

$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 5000
$toolTip.InitialDelay = 1000
$toolTip.ReshowDelay = 500

function Test-Requirements {
    $requirementsMet = $true
    $messages = @()
    
    # Check TagLib installation
    $tagLibModule = Get-Module -Name TagLibCli -ListAvailable
    if (-not $tagLibModule) {
        $requirementsMet = $false
        $messages += "TagLibCli module not found. Please install it with: Install-Module -Name TagLibCli -Force"
    } else {
        # Check if DLL exists
        $moduleDir = Split-Path $tagLibModule.Path -Parent
        $dllPath = Join-Path $moduleDir "TagLibSharp.dll"
        if (-not (Test-Path $dllPath)) {
            $requirementsMet = $false
            $messages += "TagLibSharp.dll not found in module directory. Comments functionality will be disabled"
        }
    }
    
    # Show messages in tray
    foreach ($message in $messages) {
        Show-TrayNotification -Title "Requirements Check" -Message $message -Type "Info"
    }
    
    return $requirementsMet
}

function CreateControls {
    $gap = [int]($global:fontSize * 0.8)
    $btnH = [int]($global:fontSize * 2.2)
    $btnW = 160 + $global:fontSize*2
    $leftPanelWidth = $btnW + $gap * 2
    $y = $gap
    $x = $gap



    $btnDeleteW = [int](($btnW - $gap) / 2)

    $controls.DeleteBtn = New-Object Windows.Forms.Button
    $controls.DeleteBtn.Text = "Delete"
    $controls.DeleteBtn.Enabled = $false
    $form.Controls.Add($controls.DeleteBtn)
    $controls.DeleteBtn.SetBounds($x, $y, $btnDeleteW, $btnH)

    $controls.BinBtn = New-Object Windows.Forms.Button
    $controls.BinBtn.Text = "Bin"
    $controls.BinBtn.Enabled = $false
    $form.Controls.Add($controls.BinBtn)
    $controls.BinBtn.SetBounds($x + $btnDeleteW + $gap, $y, $btnDeleteW, $btnH)
    $y += $btnH + $gap

    $controls.SortNameRadio = New-Object Windows.Forms.RadioButton
    $controls.SortNameRadio.Text = "Name"
    $controls.SortNameRadio.AutoSize = $false
    $form.Controls.Add($controls.SortNameRadio)
    $controls.SortNameRadio.SetBounds($x, $y, $btnW, $btnH)
    $y += $btnH + $gap

    $controls.SortSizeRadio = New-Object Windows.Forms.RadioButton
    $controls.SortSizeRadio.Text = "Size"
    $controls.SortSizeRadio.AutoSize = $false
    $form.Controls.Add($controls.SortSizeRadio)
    $controls.SortSizeRadio.SetBounds($x, $y, $btnW, $btnH)
    $y += $btnH + $gap

    $controls.SortCreatedRadio = New-Object Windows.Forms.RadioButton
    $controls.SortCreatedRadio.Text = "Created"
    $controls.SortCreatedRadio.AutoSize = $false
    $controls.SortCreatedRadio.Checked = $true
    $form.Controls.Add($controls.SortCreatedRadio)
    $controls.SortCreatedRadio.SetBounds($x, $y, $btnW, $btnH)
    $y += $btnH + $gap

    $controls.SelectFolder = New-Object Windows.Forms.Button
    $controls.SelectFolder.Text = "Folder"
    $form.Controls.Add($controls.SelectFolder)
    $controls.SelectFolder.SetBounds($x, $y, $btnW, $btnH)
    $y += $btnH + $gap

     if ($global:commentsEnabled) {
         $controls.CommentsBox = New-Object Windows.Forms.RichTextBox
         $controls.CommentsBox.Multiline = $true
         $controls.CommentsBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
         $controls.CommentsBox.ReadOnly = $false
         $controls.CommentsBox.Text = ""
         $form.Controls.Add($controls.CommentsBox)
         $controls.CommentsBox.SetBounds($x, $y, $btnW, 150)
         $y += 150 + $gap
 
         $controls.SaveCommentsBtn = New-Object Windows.Forms.Button
         $controls.SaveCommentsBtn.Text = "Save"
         $controls.SaveCommentsBtn.Enabled = $false
         $form.Controls.Add($controls.SaveCommentsBtn)
         $controls.SaveCommentsBtn.SetBounds($x, $y, $btnW, $btnH)
         $y += $btnH + $gap
 
         $controls.UpdateCommentsBtn = New-Object Windows.Forms.Button
         $controls.UpdateCommentsBtn.Text = "Update"
         $controls.UpdateCommentsBtn.Enabled = $true
         $controls.UpdateCommentsBtn.Visible = $true
         $form.Controls.Add($controls.UpdateCommentsBtn)
         $controls.UpdateCommentsBtn.SetBounds($x, $y, $btnW, $btnH)
         $y += $btnH + $gap
     }

    $controls.StatusStrip = New-Object Windows.Forms.StatusStrip
    
    $controls.AboutLink = New-Object Windows.Forms.ToolStripStatusLabel
    $controls.AboutLink.Text = "about"
    $controls.AboutLink.IsLink = $true
    $controls.AboutLink.LinkColor = [System.Drawing.Color]::Blue
    $controls.AboutLink.VisitedLinkColor = [System.Drawing.Color]::Purple
    $controls.StatusStrip.Items.Add($controls.AboutLink)
    
    $controls.StatusLabel = New-Object Windows.Forms.ToolStripStatusLabel
    $controls.StatusLabel.Text = "Total files: 0 | Total size: 0 MB"
    $controls.StatusLabel.Spring = $true
    $controls.StatusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
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
    $controls.ListView.Left = $leftPanelWidth
    $controls.ListView.Top = 0
    $controls.ListView.Width = $form.ClientSize.Width - $leftPanelWidth
    $controls.ListView.Height = $form.ClientSize.Height - $controls.StatusStrip.Height
    $controls.ListView.Columns.Add("File Name", -1) | Out-Null
    $controls.ListView.Columns.Add("MB", 100) | Out-Null
    $controls.ListView.Columns.Add("Created", 100) | Out-Null
    
    if ($global:commentsEnabled) {
        $controls.ListView.Columns.Add("Comments", 500) | Out-Null
    }

    $selectFolderTooltip = @"
Allows you to select another folder to display and work with its files.
"@
    $deleteTooltip = @"
Permanently deletes the selected files.
"@
    $binTooltip = @"
Moves the selected files to the Recycle Bin.
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
    $toolTip.SetToolTip($controls.BinBtn, $binTooltip.Trim())
    $toolTip.SetToolTip($controls.SortNameRadio, $sortNameTooltip.Trim())
    $toolTip.SetToolTip($controls.SortSizeRadio, $sortSizeTooltip.Trim())
    $toolTip.SetToolTip($controls.SortCreatedRadio, $sortCreatedTooltip.Trim())

    if ($global:commentsEnabled) {
        $updateTooltip = @"
Loads metadata comments for visible files in the list.
Comments are loaded on-demand to improve performance.
"@
        $toolTip.SetToolTip($controls.UpdateCommentsBtn, $updateTooltip.Trim())
    }
}

function LayoutOnlyFonts {
    $gap = [int]($global:fontSize * 0.8)
    $btnH = [int]($global:fontSize * 2.2)
    $btnW = 160 + $global:fontSize*2
    $leftPanelWidth = $btnW + $gap * 2
    $y = $gap
    $x = $gap



    $btnDeleteW = [int](($btnW - $gap) / 2)
    $controls.DeleteBtn.SetBounds($x, $y, $btnDeleteW, $btnH)
    $controls.BinBtn.SetBounds($x + $btnDeleteW + $gap, $y, $btnDeleteW, $btnH)
    $y += $btnH + $gap

    $controls.SortNameRadio.SetBounds($x, $y, $btnW, $btnH)
    $y += $btnH + $gap

    $controls.SortSizeRadio.SetBounds($x, $y, $btnW, $btnH)
    $y += $btnH + $gap

    $controls.SortCreatedRadio.SetBounds($x, $y, $btnW, $btnH)
    $y += $btnH + $gap

    $controls.SelectFolder.SetBounds($x, $y, $btnW, $btnH)
    $y += $btnH + $gap

    if ($global:commentsEnabled) {
        $controls.CommentsBox.SetBounds($x, $y, $btnW, 150)
        $y += 150 + $gap
        $controls.SaveCommentsBtn.SetBounds($x, $y, $btnW, $btnH)
        $y += $btnH + $gap
        $controls.UpdateCommentsBtn.SetBounds($x, $y, $btnW, $btnH)
    }

    $controls.ListView.Left = $leftPanelWidth
    $controls.ListView.Top = 0
    $controls.ListView.Width = $form.ClientSize.Width - $leftPanelWidth
    $controls.ListView.Height = $form.ClientSize.Height - $controls.StatusStrip.Height

    $controls.DeleteBtn.Font = $font
    $controls.BinBtn.Font = $font
    $controls.SortNameRadio.Font = $font
    $controls.SortSizeRadio.Font = $font
    $controls.SortCreatedRadio.Font = $font
    $controls.SelectFolder.Font = $font
    if ($global:commentsEnabled) {
        $controls.CommentsBox.Font = $font
        $controls.SaveCommentsBtn.Font = $font
        $controls.UpdateCommentsBtn.Font = $font
    }
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
    
    $showComments = $global:commentsEnabled
    
    foreach ($file in $global:filteredTable) {
        $displayName = $file.Name
        $item = New-Object Windows.Forms.ListViewItem($displayName)
        $item.UseItemStyleForSubItems = $false
        $item.SubItems.Add("$($file.SizeMB)")
        $displayDate = Format-ExtractedDate $file.DisplayDate $true
        $item.SubItems.Add($displayDate)
        
        if ($showComments) {
            if ($file.CommentsLoaded) {
                $comments = $file.Comments
                if ($null -eq $comments) { $comments = "" }
            } else {
                $comments = "not updated"
            }
            $item.SubItems.Add($comments)
            
            if ($file.CommentsLoaded) {
                $item.SubItems[3].ForeColor = [System.Drawing.Color]::Black
            } else {
                $item.SubItems[3].ForeColor = [System.Drawing.Color]::LightGray
            }
        }
        
        $controls.ListView.Items.Add($item) | Out-Null
    }
    $controls.DeleteBtn.Enabled = $controls.ListView.Items.Count -gt 0 -and $controls.ListView.SelectedItems.Count -gt 0
    $controls.BinBtn.Enabled = $controls.ListView.Items.Count -gt 0 -and $controls.ListView.SelectedItems.Count -gt 0
    Update-InfoLabels
    Update-ListViewTextColors
}

function Update-ListViewPreserveScroll {
    $topItemIndex = -1
    if ($null -ne $controls.ListView.TopItem) {
        $topItemIndex = $controls.ListView.TopItem.Index
    }
    
    $selectedIndexes = @()
    foreach ($item in $controls.ListView.SelectedItems) {
        $selectedIndexes += $item.Index
    }
    
    $controls.ListView.BeginUpdate()
    
    $controls.ListView.Items.Clear()
    
    $showComments = $global:commentsEnabled
    
    foreach ($file in $global:filteredTable) {
        $displayName = $file.Name
        $item = New-Object Windows.Forms.ListViewItem($displayName)
        $item.UseItemStyleForSubItems = $false
        $item.SubItems.Add("$($file.SizeMB)")
        $displayDate = Format-ExtractedDate $file.DisplayDate $true
        $item.SubItems.Add($displayDate)
        
        if ($showComments) {
            if ($file.CommentsLoaded) {
                $comments = $file.Comments
                if ($null -eq $comments) { $comments = "" }
            } else {
                $comments = "not updated"
            }
            $item.SubItems.Add($comments)
            
            if ($file.CommentsLoaded) {
                $item.SubItems[3].ForeColor = [System.Drawing.Color]::Black
            } else {
                $item.SubItems[3].ForeColor = [System.Drawing.Color]::LightGray
            }
        }
        
        $controls.ListView.Items.Add($item) | Out-Null
    }
    
    if ($topItemIndex -ge 0 -and $topItemIndex -lt $controls.ListView.Items.Count) {
        $controls.ListView.TopItem = $controls.ListView.Items[$topItemIndex]
    }
    
    foreach ($index in $selectedIndexes) {
        if ($index -lt $controls.ListView.Items.Count) {
            $controls.ListView.Items[$index].Selected = $true
        }
    }
    
    $controls.ListView.EndUpdate()
    
    $controls.DeleteBtn.Enabled = $controls.ListView.Items.Count -gt 0 -and $controls.ListView.SelectedItems.Count -gt 0
    $controls.BinBtn.Enabled = $controls.ListView.Items.Count -gt 0 -and $controls.ListView.SelectedItems.Count -gt 0
    Update-InfoLabels
    Update-ListViewTextColors
}

function Update-ListViewTextColors {
    $showComments = $global:commentsEnabled
    
    foreach ($item in $controls.ListView.Items) {
        $fileIndex = $item.Index
        if ($fileIndex -lt $global:filteredTable.Count) {
            $file = $global:filteredTable[$fileIndex]
            
            $item.Text = $file.Name
            $item.ForeColor = [System.Drawing.Color]::Black
            
            $item.SubItems[2].Text = Format-ExtractedDate $file.DisplayDate $true
            $item.SubItems[2].ForeColor = [System.Drawing.Color]::Black
            
            $item.SubItems[1].ForeColor = [System.Drawing.Color]::Black
            
            if ($showComments -and $item.SubItems.Count -gt 3) {
                if ($file.CommentsLoaded) {
                    $comments = $file.Comments
                    if ($null -eq $comments) { $comments = "" }
                } else {
                    $comments = "not updated"
                }
                $item.SubItems[3].Text = $comments
                
                if ($file.CommentsLoaded) {
                    $item.SubItems[3].ForeColor = [System.Drawing.Color]::Black
                } else {
                    $item.SubItems[3].ForeColor = [System.Drawing.Color]::LightGray
                }
            }
        }
    }
    
    if ($global:commentsEnabled) {
        $controls.ListView.AutoResizeColumn(0, [System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
        $controls.ListView.AutoResizeColumn(1, [System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
        $controls.ListView.AutoResizeColumn(2, [System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
        
        if ($showComments -and $controls.ListView.Columns.Count -gt 3) {
            $controls.ListView.Columns[3].Width = 500
        }
    } else {
        $controls.ListView.AutoResizeColumn(0, [System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
        $controls.ListView.AutoResizeColumn(2, [System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
    }
}

function Load-CommentsForVisibleItems {
    if (-not $global:commentsEnabled) {
        return
    }
    
    $topItemIndex = -1
    if ($null -ne $controls.ListView.TopItem) {
        $topItemIndex = $controls.ListView.TopItem.Index
    }
    
    $visibleCount = $controls.ListView.VisibleCount
    $startIndex = $topItemIndex
    $endIndex = [math]::Min($global:filteredTable.Count - 1, $topItemIndex + $visibleCount - 1)
    
    $endIndex = [math]::Min($global:filteredTable.Count - 1, $endIndex + 25)
    
    Write-Host "=== Loading Comments for Visible Items ===" -ForegroundColor Cyan
    Write-Host "Top item index: $topItemIndex, Visible count: $visibleCount" -ForegroundColor Yellow
    Write-Host "Range: $startIndex to $endIndex (with buffer)" -ForegroundColor Yellow
    Write-Host "Total files in filtered table: $($global:filteredTable.Count)" -ForegroundColor Yellow
    
    $loadedCount = 0
    $skippedCount = 0
    for ($i = $startIndex; $i -le $endIndex; $i++) {
        if ($i -lt $global:filteredTable.Count) {
            $file = $global:filteredTable[$i]
            
            if (-not $file.CommentsLoaded) {
                if ($i -lt $controls.ListView.Items.Count) {
                    $item = $controls.ListView.Items[$i]
                    $item.BackColor = [System.Drawing.Color]::LightGray
                    foreach ($subItem in $item.SubItems) {
                        $subItem.BackColor = [System.Drawing.Color]::LightGray
                    }
                    [System.Windows.Forms.Application]::DoEvents()
                }
                
                Write-Host "Loading comments for: $($file.Name) (index $i)" -ForegroundColor Green
                $file.Comments = Read-FileComments $file.Path
                $file.CommentsLoaded = $true
                $loadedCount++

                if ($i -lt $controls.ListView.Items.Count) {
                    $item = $controls.ListView.Items[$i]
                    if ($item.SubItems.Count -gt 3) {
                        $item.SubItems[3].Text = $file.Comments
                        $item.SubItems[3].ForeColor = [System.Drawing.Color]::Black
                    }
                    [System.Windows.Forms.Application]::DoEvents()
                }

                if ($i -lt $controls.ListView.Items.Count) {
                    $item = $controls.ListView.Items[$i]
                    $item.BackColor = [System.Drawing.Color]::White
                    foreach ($subItem in $item.SubItems) {
                        $subItem.BackColor = [System.Drawing.Color]::White
                    }
                    [System.Windows.Forms.Application]::DoEvents()
                }
            } else {
                Write-Host "Skipping already loaded: $($file.Name) (index $i)" -ForegroundColor Gray
                $skippedCount++
            }
        }
    }
    Write-Host "Loaded: $loadedCount files, Skipped: $skippedCount files" -ForegroundColor Cyan
    Write-Host "=== End Loading Comments ===" -ForegroundColor Cyan
    
    if ($loadedCount -gt 0) {
        Update-ListViewTextColors
    }
    
    # Disable the Update button after loading is completed
    if ($global:commentsEnabled -and $controls.UpdateCommentsBtn) {
        $controls.UpdateCommentsBtn.Enabled = $false
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
    $extension = [System.IO.Path]::GetExtension($fileName).ToLower()
    
    # Special handling for MP3 files with parentheses pattern
    if ($extension -eq ".mp3" -and $nameWithoutExt -match "(.+?)\((.+?)\)") {
        $beforeParentheses = $matches[1]
        $insideParentheses = $matches[2]
        
        # If content inside parentheses equals content before parentheses, return only the content inside
        if ($beforeParentheses -eq $insideParentheses) {
            return $insideParentheses
        }
    }
    
    # Original logic for other cases
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

function Read-FileComments($filePath) {
    try {
        $moduleDir = Split-Path (Get-Module -Name TagLibCli -ListAvailable).Path -Parent
        $dllPath = Join-Path $moduleDir "TagLibSharp.dll"

        if (Test-Path $dllPath) {
            Add-Type -Path $dllPath
            
            $tagFile = [TagLib.File]::Create($filePath)
            
            if ($tagFile) {
                $tags = $tagFile.Tag
                
                if ($tags) {
                    $comments = $tags.Comment
                    $tagFile.Dispose()
                    return $comments
                }
                
                $tagFile.Dispose()
            }
        }
    } catch {
        Write-Host "Error reading comments: $($_.Exception.Message)" -ForegroundColor Red
    }
    return ""
}

function Write-FileComments($filePath, $comments) {
    try {
        $moduleDir = Split-Path (Get-Module -Name TagLibCli -ListAvailable).Path -Parent
        $dllPath = Join-Path $moduleDir "TagLibSharp.dll"
        
        if (Test-Path $dllPath) {
            Add-Type -Path $dllPath
            
            $tagFile = [TagLib.File]::Create($filePath)
            
            if ($tagFile) {
                $tags = $tagFile.Tag
                
                if ($tags) {
                    $tags.Comment = $comments
                    $tagFile.Save()
                    $tagFile.Dispose()
                    return $true
                }
                
                $tagFile.Dispose()
            }
        }
    } catch {
        Write-Host "Error writing comments: $($_.Exception.Message)" -ForegroundColor Red
    }
    return $false
}

function Update-CommentsDisplay {
    if (-not $global:commentsEnabled) {
        Write-Host "Update-CommentsDisplay: Comments not enabled" -ForegroundColor Red
        return
    }
    
    $selectedCount = $controls.ListView.SelectedItems.Count
    Write-Host "Update-CommentsDisplay: Selected count = $selectedCount" -ForegroundColor Yellow
    
    if ($selectedCount -eq 1) {
        $index = $controls.ListView.SelectedItems[0].Index
        Write-Host "Update-CommentsDisplay: Selected index = $index" -ForegroundColor Yellow
        
        if ($index -lt $global:filteredTable.Count) {
            $file = $global:filteredTable[$index]
            Write-Host "Update-CommentsDisplay: File = $($file.Name), CommentsLoaded = $($file.CommentsLoaded)" -ForegroundColor Yellow
            
            if ($file -and $file.Path -and (Test-Path $file.Path)) {
                $extension = [System.IO.Path]::GetExtension($file.Path).ToLower()
                Write-Host "Update-CommentsDisplay: Extension = $extension" -ForegroundColor Yellow
                
                if ($extension -match "\.(m4a|mp3|ogg)$") {
                    $global:currentSelectedFile = $file
                    
                    if (-not $file.CommentsLoaded) {
                        Write-Host "Update-CommentsDisplay: Loading comments for $($file.Name)" -ForegroundColor Cyan
                        $file.Comments = Read-FileComments $file.Path
                        $file.CommentsLoaded = $true
                        Write-Host "Update-CommentsDisplay: Loaded comments = '$($file.Comments)'" -ForegroundColor Green
                        
                        Update-ListViewTextColors
                    } else {
                        Write-Host "Update-CommentsDisplay: Using cached comments = '$($file.Comments)'" -ForegroundColor Green
                    }
                    
                    $comments = $file.Comments
                    $global:originalCommentsText = $comments
                    $controls.CommentsBox.Text = $comments
                    $controls.SaveCommentsBtn.Enabled = $false
                } else {
                    Write-Host "Update-CommentsDisplay: Not an audio file" -ForegroundColor Red
                    $global:currentSelectedFile = $null
                    $global:originalCommentsText = ""
                    $controls.CommentsBox.Text = ""
                    $controls.SaveCommentsBtn.Enabled = $false
                }
            } else {
                Write-Host "Update-CommentsDisplay: File or path invalid" -ForegroundColor Red
                $global:currentSelectedFile = $null
                $global:originalCommentsText = ""
                $controls.CommentsBox.Text = ""
                $controls.SaveCommentsBtn.Enabled = $false
            }
        } else {
            Write-Host "Update-CommentsDisplay: Index out of range" -ForegroundColor Red
            $global:currentSelectedFile = $null
            $global:originalCommentsText = ""
            $controls.CommentsBox.Text = ""
            $controls.SaveCommentsBtn.Enabled = $false
        }
    } else {
        Write-Host "Update-CommentsDisplay: Multiple or no selection, clearing" -ForegroundColor Yellow
        $global:currentSelectedFile = $null
        $global:originalCommentsText = ""
        $controls.CommentsBox.Text = ""
        $controls.SaveCommentsBtn.Enabled = $false
    }
}

 

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

        # Use a timer to dispose the notification asynchronously instead of blocking
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = $Duration

        # Capture variables in closure to avoid scope issues
        $localNotifyIcon = $notifyIcon
        $localTimer = $timer

        $timer.Add_Tick({
            try {
                if ($null -ne $localNotifyIcon -and -not $localNotifyIcon.IsDisposed) {
                    $localNotifyIcon.Dispose()
                }
            } catch {
                # Ignore disposal errors
            }
            try {
                if ($null -ne $localTimer -and -not $localTimer.IsDisposed) {
                    $localTimer.Dispose()
                }
            } catch {
                # Ignore disposal errors
            }
        })
        $timer.Start()

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
            
            $fileObj = [PSCustomObject]@{
                Name   = $displayName
                SizeMB = [math]::Round($file.Length / 1MB, 2)
                Path   = $file.FullName
                CreationTime = $file.CreationTime
                ExtractedDate = $extractedDate
                DisplayDate = $displayDate
                OrigName = $file.Name
                Comments = $null
                CommentsLoaded = $false
            }
            
            $global:fileTable += $fileObj
        }
        
        $controls.SortCreatedRadio.Checked = $true
        $global:fileTable = $global:fileTable | Sort-Object DisplayDate -Descending
        Invoke-Search
    } else {
        $global:fileTable = @()
        $controls.SortCreatedRadio.Checked = $true
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
        $hasSelection = $controls.ListView.SelectedItems.Count -gt 0
        $controls.DeleteBtn.Enabled = $hasSelection
        $controls.BinBtn.Enabled = $hasSelection
        Update-InfoLabels
        Update-CommentsDisplay
    })
    
    $controls.SortNameRadio.Add_CheckedChanged({
        if ($controls.SortNameRadio.Checked) {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
           $global:filteredTable = $global:filteredTable | Sort-Object @{Expression="Name"; Ascending=$true}, @{Expression="DisplayDate"; Ascending=$false}
            Update-ListView
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })
    $controls.SortSizeRadio.Add_CheckedChanged({
        if ($controls.SortSizeRadio.Checked) {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $global:filteredTable = $global:filteredTable | Sort-Object SizeMB -Descending
            Update-ListView
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })
    $controls.SortCreatedRadio.Add_CheckedChanged({
        if ($controls.SortCreatedRadio.Checked) {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $global:filteredTable = $global:filteredTable | Sort-Object DisplayDate -Descending
            Update-ListView
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })
    $controls.DeleteBtn.Add_Click({
        $toDeleteIndexes = @()
        foreach ($item in $controls.ListView.SelectedItems) {
            $toDeleteIndexes += $item.Index
        }
        $toDeleteIndexes = $toDeleteIndexes | Sort-Object -Descending
        $deleted = 0
        foreach ($i in $toDeleteIndexes) {
            $file = $global:filteredTable[$i]
            try {
                Remove-Item -Path $file.Path -Force
                $deleted++
                $global:fileTable = $global:fileTable | Where-Object { $_.Path -ne $file.Path }
                $global:filteredTable = $global:filteredTable | Where-Object { $_.Path -ne $file.Path }
            } catch {
            }
        }
        Update-ListViewPreserveScroll
        Show-TrayNotification -Title "Done" -Message "$deleted file(s) permanently deleted."
    })
    $controls.BinBtn.Add_Click({
        $toDeleteIndexes = @()
        foreach ($item in $controls.ListView.SelectedItems) {
            $toDeleteIndexes += $item.Index
        }
        $toDeleteIndexes = $toDeleteIndexes | Sort-Object -Descending
        $deleted = 0
        foreach ($i in $toDeleteIndexes) {
            $file = $global:filteredTable[$i]
            try {
                $success = Move-FileToRecycleBin $file.Path
                if ($success) {
                    $deleted++
                    $global:fileTable = $global:fileTable | Where-Object { $_.Path -ne $file.Path }
                    $global:filteredTable = $global:filteredTable | Where-Object { $_.Path -ne $file.Path }
                }
            } catch {
            }
        }
        Update-ListViewPreserveScroll
        Show-TrayNotification -Title "Done" -Message "$deleted file(s) moved to Recycle Bin."
    })
    $controls.ListView.Add_DoubleClick({
        if ($controls.ListView.SelectedItems.Count -eq 1) {
            $index = $controls.ListView.SelectedItems[0].Index
            $file = $global:filteredTable[$index]
            try {
                # Use Start-Process for PowerShell 7 compatibility
                Start-Process -FilePath $file.Path -ErrorAction Stop
            } catch {
                Show-TrayNotification -Title "Error" -Message "Cannot open file: $($file.Path)" -Type "Error"
            }
        }
    })
    
    # Track scroll position for auto-updating comments
    $global:lastScrollTop = 0
    $global:scrollTimer = $null

    # Create a timer to handle scroll events
    $global:scrollTimer = New-Object System.Windows.Forms.Timer
    $global:scrollTimer.Interval = 500  # 500ms delay
    $global:scrollTimer.Add_Tick({
        $currentTop = $controls.ListView.TopItem
        if ($currentTop -and $currentTop.Index -ne $global:lastScrollTop) {
            $global:lastScrollTop = $currentTop.Index
            if ($global:commentsEnabled) {
                # Enable the Update button when scrolling occurs
                $controls.UpdateCommentsBtn.Enabled = $true
                $controls.UpdateCommentsBtn.PerformClick()
            }
        }
        $global:scrollTimer.Stop()
    })

    # Add mouse wheel event handler for scroll detection
    $controls.ListView.Add_MouseWheel({
        if ($global:commentsEnabled) {
            # Enable the Update button when scrolling occurs
            $controls.UpdateCommentsBtn.Enabled = $true
            $global:scrollTimer.Stop()
            $global:scrollTimer.Start()
        }
    })

    # Add mouse capture changed event handler to detect scrollbar dragging
    $controls.ListView.Add_MouseCaptureChanged({
        if ($global:commentsEnabled) {
            $currentTop = $controls.ListView.TopItem
            $currentIndex = if ($currentTop) { $currentTop.Index } else { -1 }
            if ($currentTop -and $currentIndex -ne $global:lastScrollTop) {
                # Enable the Update button only when scroll position changed
                $controls.UpdateCommentsBtn.Enabled = $true
                $global:scrollTimer.Stop()
                $global:scrollTimer.Start()
            }
        }
    })


    if ($global:commentsEnabled) {
        $controls.SaveCommentsBtn.Add_Click({
            if ($global:currentSelectedFile -and $controls.CommentsBox.Text -ne $global:originalCommentsText) {
                $newComments = $controls.CommentsBox.Text
                
                if (-not [string]::IsNullOrWhiteSpace($newComments)) {
                    $controls.SaveCommentsBtn.Enabled = $false
                    $originalText = $controls.SaveCommentsBtn.Text
                    $controls.SaveCommentsBtn.Text = "Saving..."
                    
                    [System.Windows.Forms.Application]::DoEvents()
                    
                    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
                    
                    Write-Host "=== Saving Comments Started ===" -ForegroundColor Cyan
                    $success = Write-FileComments $global:currentSelectedFile.Path $newComments
                    Write-Host "=== Saving Comments Completed ===" -ForegroundColor Cyan
                    
                     $controls.SaveCommentsBtn.Text = $originalText
                     $form.Cursor = [System.Windows.Forms.Cursors]::Default
                     
                     [System.Windows.Forms.Application]::DoEvents()
                     
                     if ($success) {
                         $fileName = [System.IO.Path]::GetFileName($global:currentSelectedFile.Path)
                         Show-TrayNotification -Title "Comments Updated" -Message "Comments for '$fileName' successfully saved."
                         $global:originalCommentsText = $newComments
                         
                         $global:currentSelectedFile.Comments = $newComments
                         $global:currentSelectedFile.CommentsLoaded = $true
                         
                         Update-ListViewTextColors
                         
                         $controls.SaveCommentsBtn.Enabled = $false
                     } else {
                        $fileName = [System.IO.Path]::GetFileName($global:currentSelectedFile.Path)
                        Show-TrayNotification -Title "Error" -Message "Failed to save comments for '$fileName'" -Type "Error"
                    }
                } else {
                    Show-TrayNotification -Title "Error" -Message "Cannot save empty comments" -Type "Error"
                }
            }
        })

        $controls.UpdateCommentsBtn.Add_Click({
            $controls.UpdateCommentsBtn.Enabled = $false
            $originalText = $controls.UpdateCommentsBtn.Text
            $controls.UpdateCommentsBtn.Text = "Loading..."
            
            [System.Windows.Forms.Application]::DoEvents()
            
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            
            Write-Host "=== Manual Comments Loading Started ===" -ForegroundColor Cyan
            Load-CommentsForVisibleItems
            Write-Host "=== Manual Comments Loading Completed ===" -ForegroundColor Cyan
            
            $controls.UpdateCommentsBtn.Text = $originalText
            $controls.UpdateCommentsBtn.Enabled = $false  # Disable after update
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            
            [System.Windows.Forms.Application]::DoEvents()
        })

        $controls.CommentsBox.Add_TextChanged({
            if ($global:currentSelectedFile) {
                $currentText = $controls.CommentsBox.Text
                $originalText = $global:originalCommentsText
                
                $isDifferent = $currentText -ne $originalText
                $isNotEmpty = -not [string]::IsNullOrWhiteSpace($currentText)
                
                $controls.SaveCommentsBtn.Enabled = $isDifferent -and $isNotEmpty
            }
        })
    }
    
    $controls.AboutLink.Add_Click({
        Start-Process "https://github.com/SVDotsenko/cleaner/blob/main/readme.md"
    })
}

$form.Add_Resize({
    if ($null -ne $controls.ListView) {
        $gap = [int]($global:fontSize * 0.8)
        $btnW = 160 + $global:fontSize*2
        $leftPanelWidth = $btnW + $gap * 2
        $controls.ListView.Left = $leftPanelWidth
        $controls.ListView.Width = $form.ClientSize.Width - $leftPanelWidth
        $controls.ListView.Height = $form.ClientSize.Height - $controls.StatusStrip.Height
        
        if ($global:commentsEnabled) {
            $controls.ListView.AutoResizeColumn(0, [System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
            $controls.ListView.AutoResizeColumn(1, [System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
            $controls.ListView.AutoResizeColumn(2, [System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
            
            if ($controls.ListView.Columns.Count -gt 3) {
                $controls.ListView.Columns[3].Width = 500
            }
        } else {
            $controls.ListView.AutoResizeColumn(0, [System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
        }
    }
})

$form.Topmost = $false

$form.Add_Shown({
    $global:commentsEnabled = Test-Requirements
    
    CreateControls
    Set-AllFonts $global:fontSize
    BindHandlers
    Get-FilesFromFolder
    
    # Initialize baseline top index to avoid false scroll detection on first clicks
    if ($null -ne $controls.ListView.TopItem) {
        $global:lastScrollTop = $controls.ListView.TopItem.Index
    } else {
        $global:lastScrollTop = 0
    }
    
    # Auto-load comments for visible items on startup
    if ($global:commentsEnabled) {
        Write-Host "=== Auto-loading comments on startup ===" -ForegroundColor Cyan
        $controls.UpdateCommentsBtn.PerformClick()
        Write-Host "=== Auto-loading completed ===" -ForegroundColor Cyan
    }
    
    $form.Activate()
})

[void]$form.ShowDialog()
