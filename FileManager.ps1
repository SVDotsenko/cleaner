Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

$global:folderPath = "G:\My Drive\recordings"
$global:fontSize = 14
$global:fontFamily = "Segoe UI"
$global:fileTable = @()
$global:filteredTable = @()
$global:activeSortButton = $null
$global:maxColumnWidths = @{ Name = 0; Date = 0; Duration = 0 }
$global:currentSelectedFile = $null
$global:originalCommentsText = ""
$global:lastScrollTop = 0
$global:scrollTimer = $null
$global:selectedYears = @()  # Array of selected years for filtering
$global:backgroundTimer = $null  # Timer for async comment loading
$global:isBackgroundLoading = $false  # Flag to track background loading state
$global:backgroundFileIndexes = @()  # Array of file indexes to process
$global:backgroundCurrentIndex = 0  # Current index being processed
$global:commentsEnabled = $true  # Add this variable for background loading

function Test-Requirements {
    # Check TagLib installation
    $tagLibModule = Get-Module -Name TagLibCli -ListAvailable
    if (-not $tagLibModule) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "TagLib is required for this application to work.`n`nPlease install TagLib by following the instructions at:`nhttps://github.com/SVDotsenko/cleaner/blob/after-youtube/readme.md`n`nClick OK to open the installation guide.",
            "TagLib Required",
            [System.Windows.Forms.MessageBoxButtons]::OKCancel,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            Start-Process "https://github.com/SVDotsenko/cleaner/blob/after-youtube/readme.md"
        }
        return $false
    }

    # Check if DLL exists
    $moduleDir = Split-Path $tagLibModule.Path -Parent
    $dllPath = Join-Path $moduleDir "TagLibSharp.dll"
    if (-not (Test-Path $dllPath)) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "TagLib is required for this application to work.`n`nPlease install TagLib by following the instructions at:`nhttps://github.com/SVDotsenko/cleaner/blob/after-youtube/readme.md`n`nClick OK to open the installation guide.",
            "TagLib Required",
            [System.Windows.Forms.MessageBoxButtons]::OKCancel,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            Start-Process "https://github.com/SVDotsenko/cleaner/blob/after-youtube/readme.md"
        }
        return $false
    }

    return $true
}

# Check requirements at startup - exit if not met
if (-not (Test-Requirements)) {
    exit
}

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

    # Add Sort GroupBox
    $sortGroupHeight = $btnH * 3 + $gap * 2 + 30  # Increased padding from 20 to 30
    $controls.SortGroupBox = New-Object Windows.Forms.GroupBox
    $controls.SortGroupBox.Text = "Sort"
    $controls.SortGroupBox.SetBounds($x, $y, $btnW, $sortGroupHeight)
    $form.Controls.Add($controls.SortGroupBox)

    $radioY = 25  # Increased from 20 to 25
    $controls.SortNameRadio = New-Object Windows.Forms.RadioButton
    $controls.SortNameRadio.Text = "Name"
    $controls.SortNameRadio.AutoSize = $false
    $controls.SortGroupBox.Controls.Add($controls.SortNameRadio)
    $controls.SortNameRadio.SetBounds(10, $radioY, $btnW - 20, $btnH)
    $radioY += $btnH + $gap

    $controls.SortSizeRadio = New-Object Windows.Forms.RadioButton
    $controls.SortSizeRadio.Text = "Duration"
    $controls.SortSizeRadio.AutoSize = $false
    $controls.SortGroupBox.Controls.Add($controls.SortSizeRadio)
    $controls.SortSizeRadio.SetBounds(10, $radioY, $btnW - 20, $btnH)
    $radioY += $btnH + $gap

    $controls.SortCreatedRadio = New-Object Windows.Forms.RadioButton
    $controls.SortCreatedRadio.Text = "Created"
    $controls.SortCreatedRadio.AutoSize = $false
    $controls.SortCreatedRadio.Checked = $true
    $controls.SortGroupBox.Controls.Add($controls.SortCreatedRadio)
    $controls.SortCreatedRadio.SetBounds(10, $radioY, $btnW - 20, $btnH)

    $y += $sortGroupHeight + $gap

    $controls.SelectFolder = New-Object Windows.Forms.Button
    $controls.SelectFolder.Text = "Folder"
    $form.Controls.Add($controls.SelectFolder)
    $controls.SelectFolder.SetBounds($x, $y, $btnW, $btnH)
    $y += $btnH + $gap

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

    # Add Filter GroupBox
    $filterGroupHeight = $btnH * 2 + $gap + 30  # Increased padding from 20 to 30
    $controls.FilterGroupBox = New-Object Windows.Forms.GroupBox
    $controls.FilterGroupBox.Text = "Filter"
    $controls.FilterGroupBox.SetBounds($x, $y, $btnW, $filterGroupHeight)
    $form.Controls.Add($controls.FilterGroupBox)

    $filterRadioY = 25  # Increased from 20 to 25
    $controls.ThisYearRadio = New-Object Windows.Forms.RadioButton
    $controls.ThisYearRadio.Text = "Last 20"
    $controls.ThisYearRadio.AutoSize = $false
    $controls.ThisYearRadio.Checked = $true
    $controls.FilterGroupBox.Controls.Add($controls.ThisYearRadio)
    $controls.ThisYearRadio.SetBounds(10, $filterRadioY, $btnW - 20, $btnH)
    $filterRadioY += $btnH + $gap

    $controls.AllYearsRadio = New-Object Windows.Forms.RadioButton
    $controls.AllYearsRadio.Text = "All"
    $controls.AllYearsRadio.AutoSize = $false
    $controls.FilterGroupBox.Controls.Add($controls.AllYearsRadio)
    $controls.AllYearsRadio.SetBounds(10, $filterRadioY, $btnW - 20, $btnH)

    $y += $filterGroupHeight + $gap

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

    # Add full-width ProgressBar that will overlay StatusStrip during background loading
    $controls.ProgressBar = New-Object Windows.Forms.ProgressBar
    $controls.ProgressBar.Visible = $false
    $controls.ProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $controls.ProgressBar.Step = 1
    $controls.ProgressBar.Maximum = 100
    $controls.ProgressBar.Height = 22
    $form.Controls.Add($controls.ProgressBar)

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
    $controls.ListView.Columns.Add("Duration", 100) | Out-Null
    $controls.ListView.Columns.Add("Created", 100) | Out-Null
    
    $controls.ListView.Columns.Add("Comments", 500) | Out-Null

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
Sorts the file list by duration (from longest to shortest).
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

    # Update Sort GroupBox layout
    $sortGroupHeight = $btnH * 3 + $gap * 2 + 30  # Increased padding from 20 to 30
    $controls.SortGroupBox.SetBounds($x, $y, $btnW, $sortGroupHeight)

    $radioY = 25  # Increased from 20 to 25
    $controls.SortNameRadio.SetBounds(10, $radioY, $btnW - 20, $btnH)
    $radioY += $btnH + $gap
    $controls.SortSizeRadio.SetBounds(10, $radioY, $btnW - 20, $btnH)
    $radioY += $btnH + $gap
    $controls.SortCreatedRadio.SetBounds(10, $radioY, $btnW - 20, $btnH)

    $y += $sortGroupHeight + $gap

    $controls.SelectFolder.SetBounds($x, $y, $btnW, $btnH)
    $y += $btnH + $gap

    $controls.CommentsBox.SetBounds($x, $y, $btnW, 150)
    $y += 150 + $gap
    $controls.SaveCommentsBtn.SetBounds($x, $y, $btnW, $btnH)
    $y += $btnH + $gap

    # Update Filter GroupBox layout
    $filterGroupHeight = $btnH * 2 + $gap + 30  # Increased padding from 20 to 30
    $controls.FilterGroupBox.SetBounds($x, $y, $btnW, $filterGroupHeight)

    $filterRadioY = 25  # Increased from 20 to 25
    $controls.ThisYearRadio.SetBounds(10, $filterRadioY, $btnW - 20, $btnH)
    $filterRadioY += $btnH + $gap
    $controls.AllYearsRadio.SetBounds(10, $filterRadioY, $btnW - 20, $btnH)

    $y += $filterGroupHeight + $gap

    $controls.ListView.Left = $leftPanelWidth
    $controls.ListView.Top = 0
    $controls.ListView.Width = $form.ClientSize.Width - $leftPanelWidth
    $controls.ListView.Height = $form.ClientSize.Height - $controls.StatusStrip.Height

    $controls.DeleteBtn.Font = $font
    $controls.BinBtn.Font = $font
    $controls.SortGroupBox.Font = $font
    $controls.SortNameRadio.Font = $font
    $controls.SortSizeRadio.Font = $font
    $controls.SortCreatedRadio.Font = $font
    $controls.SelectFolder.Font = $font
    $controls.CommentsBox.Font = $font
    $controls.SaveCommentsBtn.Font = $font
    $controls.FilterGroupBox.Font = $font
    $controls.ThisYearRadio.Font = $font
    $controls.AllYearsRadio.Font = $font
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
            $sum += $file.Duration
        }
        $sumFormatted = Format-Duration $sum
        $controls.StatusLabel.Text = "Selected files: $selectedCount | Selected duration: $sumFormatted"
    } else {
        $count = $global:filteredTable.Count
        $sum = 0
        if ($count -gt 0) {
            $sum = ($global:filteredTable | Measure-Object -Property Duration -Sum).Sum
        }
        $sumFormatted = Format-Duration $sum
        $yearInfo = ""
        if ($controls.ThisYearRadio.Checked) {
            $yearInfo = " | Filter: Last 20"
        }
        $controls.StatusLabel.Text = "Total files: $count | Total duration: $sumFormatted$yearInfo"
    }
}

function Format-Duration($seconds) {
    if ($null -eq $seconds -or $seconds -eq 0) { return "00:00:00" }

    # Ensure we have a valid number and convert to integer
    try {
        $secondsInt = [int][math]::Floor([double]$seconds)
        $hours = [int][math]::Floor($secondsInt / 3600)
        $minutes = [int][math]::Floor(($secondsInt % 3600) / 60)
        $secs = [int]($secondsInt % 60)
        return "{0:D2}:{1:D2}:{2:D2}" -f $hours, $minutes, $secs
    } catch {
        return "00:00:00"
    }
}

function Update-ListView {
    $controls.ListView.Items.Clear()
    
    foreach ($file in $global:filteredTable) {
        $displayName = $file.Name
        $item = New-Object Windows.Forms.ListViewItem($displayName)
        $item.UseItemStyleForSubItems = $false
        $item.SubItems.Add((Format-Duration $file.Duration))
        $displayDate = Format-ExtractedDate $file.DisplayDate $true
        $item.SubItems.Add($displayDate)
        
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
    
    foreach ($file in $global:filteredTable) {
        $displayName = $file.Name
        $item = New-Object Windows.Forms.ListViewItem($displayName)
        $item.UseItemStyleForSubItems = $false
        $item.SubItems.Add((Format-Duration $file.Duration))
        $displayDate = Format-ExtractedDate $file.DisplayDate $true
        $item.SubItems.Add($displayDate)
        
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

function Apply-YearFilter {
    if ($controls.ThisYearRadio.Checked) {
        # Show last 20 records sorted by date (newest first)
        $global:filteredTable = $global:fileTable |
            Sort-Object DisplayDate -Descending |
            Select-Object -First 20
    } else {
        # All records - show all files
        $global:filteredTable = $global:fileTable
    }
}

function Update-ListViewTextColors {
    foreach ($item in $controls.ListView.Items) {
        $fileIndex = $item.Index
        if ($fileIndex -lt $global:filteredTable.Count) {
            $file = $global:filteredTable[$fileIndex]
            
            $item.Text = $file.Name
            $item.ForeColor = [System.Drawing.Color]::Black
            
            $item.SubItems[2].Text = Format-ExtractedDate $file.DisplayDate $true
            $item.SubItems[2].ForeColor = [System.Drawing.Color]::Black
            
            $item.SubItems[1].Text = Format-Duration $file.Duration
            $item.SubItems[1].ForeColor = [System.Drawing.Color]::Black
            
            if ($item.SubItems.Count -gt 3) {
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
    
    $controls.ListView.AutoResizeColumn(0, [System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
    $controls.ListView.AutoResizeColumn(1, [System.Windows.Forms.ColumnHeaderAutoResizeStyle]::HeaderSize)
    $controls.ListView.AutoResizeColumn(2, [System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)

    if ($controls.ListView.Columns.Count -gt 3) {
        # Calculate available width for comments column with scrollbar consideration
        $totalWidth = $controls.ListView.Width
        $col0Width = $controls.ListView.Columns[0].Width
        $col1Width = $controls.ListView.Columns[1].Width
        $col2Width = $controls.ListView.Columns[2].Width

        # Account for scrollbar width (typically 17-20px) and some padding
        $scrollbarWidth = 20
        $padding = 5
        $availableWidth = $totalWidth - $col0Width - $col1Width - $col2Width - $scrollbarWidth - $padding

        if ($availableWidth -gt 50) { # Minimum width check
            $controls.ListView.Columns[3].Width = $availableWidth
        }
    }
}

function Load-CommentsForVisibleItems {
    # Ensure filtered table exists
    if (-not $global:filteredTable) { $global:filteredTable = @() }
    $totalCount = 0
    try { $totalCount = $global:filteredTable.Count } catch { $totalCount = 0 }
    
    # Nothing to do if there are no files or no rows in the ListView
    if ($totalCount -le 0 -or $controls.ListView.Items.Count -le 0) {
        return
    }

    # Determine top item index and clamp to valid range
    $topItemIndex = 0
    if ($null -ne $controls.ListView.TopItem) {
        $topItemIndex = $controls.ListView.TopItem.Index
    }
    $topItemIndex = [math]::Max(0, [math]::Min($topItemIndex, $totalCount - 1))

    # Determine visible range, always at least 1 item
    $visibleCount = $controls.ListView.VisibleCount
    if ($visibleCount -le 0) { $visibleCount = 1 }
    $startIndex = $topItemIndex
    $endIndex = [math]::Min($totalCount - 1, $topItemIndex + $visibleCount - 1)
    $endIndex = [math]::Min($totalCount - 1, $endIndex + 25)  # small buffer

    $loadedCount = 0
    $skippedCount = 0
    $itemsCount = $controls.ListView.Items.Count

    for ($i = $startIndex; $i -le $endIndex; $i++) {
        if ($i -ge 0 -and $i -lt $totalCount) {
            $file = $global:filteredTable[$i]
            if ($null -eq $file) { continue }

            if (-not $file.CommentsLoaded) {
                if ($i -lt $itemsCount) {
                    $item = $controls.ListView.Items[$i]
                    if ($null -ne $item) {
                        $item.BackColor = [System.Drawing.Color]::LightGray
                        foreach ($subItem in $item.SubItems) {
                            $subItem.BackColor = [System.Drawing.Color]::LightGray
                        }
                        [System.Windows.Forms.Application]::DoEvents()
                    }
                }

                $metadata = Read-FileMetadata $file.Path
                $file.Comments = $metadata.Comments
                $file.Duration = $metadata.Duration
                $file.CommentsLoaded = $true
                $loadedCount++

                if ($i -lt $itemsCount) {
                    $item = $controls.ListView.Items[$i]
                    if ($null -ne $item -and $item.SubItems.Count -gt 3) {
                        $item.SubItems[3].Text = $file.Comments
                        $item.SubItems[3].ForeColor = [System.Drawing.Color]::Black
                    }
                    [System.Windows.Forms.Application]::DoEvents()
                }

                if ($i -lt $itemsCount) {
                    $item = $controls.ListView.Items[$i]
                    if ($null -ne $item) {
                        $item.BackColor = [System.Drawing.Color]::White
                        foreach ($subItem in $item.SubItems) {
                            $subItem.BackColor = [System.Drawing.Color]::White
                        }
                        [System.Windows.Forms.Application]::DoEvents()
                    }
                }
            } else {
                $skippedCount++
            }
        }
    }

    if ($loadedCount -gt 0) {
        Update-ListViewTextColors
    }
    
    # Start background loading of all remaining comments
    Start-BackgroundCommentLoading
}

function Start-BackgroundCommentLoading {
    if (-not $global:commentsEnabled -or $global:isBackgroundLoading) {
        return
    }
    
    # Get count of unloaded metadata from ALL files (not just filtered)
    $unloadedCount = ($global:fileTable | Where-Object { -not $_.CommentsLoaded }).Count

    if ($unloadedCount -eq 0) {
        return
    }

    # Collect indexes of files that need comments loaded from ALL files
    $global:backgroundFileIndexes = @()
    for ($i = 0; $i -lt $global:fileTable.Count; $i++) {
        $file = $global:fileTable[$i]
        if ($null -ne $file -and -not $file.CommentsLoaded) {
            $global:backgroundFileIndexes += $i
        }
    }
    
    $global:backgroundCurrentIndex = 0
    
    # Calculate StatusStrip position and size
    $statusStripTop = $form.ClientSize.Height - $controls.StatusStrip.Height
    $statusStripLeft = 0
    $statusStripWidth = $form.ClientSize.Width

    # Calculate statistics text width (approximate)
    $statsWidth = $controls.AboutLink.Width + 20  # Add some padding

    # Position full-width ProgressBar to cover only the right part of StatusStrip
    $controls.ProgressBar.Left = $statsWidth
    $controls.ProgressBar.Top = $statusStripTop
    $controls.ProgressBar.Width = $statusStripWidth - $statsWidth
    $controls.ProgressBar.Height = $controls.StatusStrip.Height
    $controls.ProgressBar.BringToFront()

    # Show ProgressBar
    $controls.ProgressBar.Minimum = 0
    $controls.ProgressBar.Maximum = $global:backgroundFileIndexes.Count
    $controls.ProgressBar.Value = 0
    $controls.ProgressBar.Visible = $true

    # Create and configure timer
    $global:backgroundTimer = New-Object System.Windows.Forms.Timer
    $global:backgroundTimer.Interval = 50  # Reduced interval for smoother updates
    $global:backgroundTimer.Add_Tick({
        if ($global:backgroundCurrentIndex -lt $global:backgroundFileIndexes.Count) {
            $fileIndex = $global:backgroundFileIndexes[$global:backgroundCurrentIndex]
            $file = $global:fileTable[$fileIndex]
            
            Write-Host "Background loading metadata for: $($file.Name) (index $($global:backgroundCurrentIndex + 1) of $($global:backgroundFileIndexes.Count))" -ForegroundColor Green

            if ($null -ne $file -and -not $file.CommentsLoaded) {
                # Load metadata using inline function to avoid scope issues
                try {
                    $moduleDir = Split-Path (Get-Module -Name TagLibCli -ListAvailable).Path -Parent
                    $dllPath = Join-Path $moduleDir "TagLibSharp.dll"
                    
                    if (Test-Path $dllPath) {
                        Add-Type -Path $dllPath
                        $tagFile = [TagLib.File]::Create($file.Path)
                        
                        if ($tagFile) {
                            $tags = $tagFile.Tag
                            if ($tags) {
                                $file.Comments = $tags.Comment
                            } else {
                                $file.Comments = ""
                            }

                            # Get duration
                            $properties = $tagFile.Properties
                            if ($properties) {
                                $file.Duration = $properties.Duration.TotalSeconds
                            } else {
                                $file.Duration = 0
                            }

                            $tagFile.Dispose()
                        } else {
                            $file.Comments = ""
                            $file.Duration = 0
                        }
                    } else {
                        $file.Comments = ""
                        $file.Duration = 0
                    }
                } catch {
                    $file.Comments = ""
                    $file.Duration = 0
                }
                
                $file.CommentsLoaded = $true
                
                # Update progress bar
                $controls.ProgressBar.Value = $global:backgroundCurrentIndex + 1

                # Force UI update
                [System.Windows.Forms.Application]::DoEvents()
            }
            
            $global:backgroundCurrentIndex++
        } else {
            # All files processed
            # Hide progress bar
            $controls.ProgressBar.Visible = $false

            # Re-enable "All" radio button
            $controls.AllYearsRadio.Enabled = $true

            # Stop timer and reset state
            $global:backgroundTimer.Stop()
            $global:backgroundTimer.Dispose()
            $global:backgroundTimer = $null
            $global:isBackgroundLoading = $false
            $global:backgroundFileIndexes = @()
            $global:backgroundCurrentIndex = 0
            
            # Update final status
            Update-InfoLabels
        }
    })
    
    # Start background loading
    $global:isBackgroundLoading = $true
    $global:backgroundTimer.Start()
}

function Stop-BackgroundCommentLoading {
    if ($global:backgroundTimer -and $global:isBackgroundLoading) {
        $global:backgroundTimer.Stop()
        $global:backgroundTimer.Dispose()
        $global:backgroundTimer = $null
        $global:isBackgroundLoading = $false
        $global:backgroundFileIndexes = @()
        $global:backgroundCurrentIndex = 0
        
        # Hide progress bar
        $controls.ProgressBar.Visible = $false

        # Re-enable "All" radio button
        $controls.AllYearsRadio.Enabled = $true

        # Update final status
        Update-InfoLabels
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
    
    # Special handling for MP3 files
    if ($extension -eq ".mp3") {
        # Check if first character is not a letter
        if ($nameWithoutExt.Length -gt 0 -and -not [char]::IsLetter($nameWithoutExt[0])) {
            # Step 1: Parse string by "_" character
            $parts = $nameWithoutExt -split "_"
            if ($parts.Count -gt 0) {
                # Step 2: Take first element
                $firstElement = $parts[0]

                # Step 3: Remove all brackets
                $withoutBrackets = $firstElement -replace "[\(\)\[\]{}]", ""

                # Step 4: Split into 2 parts and compare
                if ($withoutBrackets.Length -gt 0) {
                    $halfLength = [math]::Floor($withoutBrackets.Length / 2)
                    if ($halfLength -gt 0 -and $withoutBrackets.Length -eq $halfLength * 2) {
                        $firstHalf = $withoutBrackets.Substring(0, $halfLength)
                        $secondHalf = $withoutBrackets.Substring($halfLength, $halfLength)

                        # If both parts are equal, return first part
                        if ($firstHalf -eq $secondHalf) {
                            return $firstHalf
                        }
                    }
                }
            }
            # If conditions not met, return full filename
            return $fileName
        }

        # Original MP3 logic for files starting with letter or with parentheses pattern
        if ($nameWithoutExt -match "(.+?)\((.+?)\)") {
            $beforeParentheses = $matches[1]
            $insideParentheses = $matches[2]

            # If content inside parentheses equals content before parentheses, return only the content inside
            if ($beforeParentheses -eq $insideParentheses) {
                return $insideParentheses
            }
        }
    }
    
    # Special handling for M4A files that don't start with a letter
    if ($extension -eq ".m4a" -and $nameWithoutExt.Length -gt 0 -and -not [char]::IsLetter($nameWithoutExt[0])) {
        if ($nameWithoutExt -match "(.+?)_") {
            return $matches[1]
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

function Read-FileMetadata($filePath) {
    try {
        $moduleDir = Split-Path (Get-Module -Name TagLibCli -ListAvailable).Path -Parent
        $dllPath = Join-Path $moduleDir "TagLibSharp.dll"

        if (Test-Path $dllPath) {
            Add-Type -Path $dllPath
            
            $tagFile = [TagLib.File]::Create($filePath)
            
            if ($tagFile) {
                $tags = $tagFile.Tag
                $properties = $tagFile.Properties

                $result = @{
                    Comments = ""
                    Duration = 0
                }

                if ($tags) {
                    $result.Comments = $tags.Comment
                    if ($null -eq $result.Comments) { $result.Comments = "" }
                }

                if ($properties) {
                    $result.Duration = $properties.Duration.TotalSeconds
                }
                
                $tagFile.Dispose()
                return $result
            }
        }
    } catch {
        Write-Host "Error reading metadata: $($_.Exception.Message)" -ForegroundColor Red
    }

    return @{
        Comments = ""
        Duration = 0
    }
}

function Write-FileMetadata($filePath, $comments) {
    try {
        $moduleDir = Split-Path (Get-Module -Name TagLibCli -ListAvailable).Path -Parent
        $dllPath = Join-Path $moduleDir "TagLibSharp.dll"

        if (Test-Path $dllPath) {
            Add-Type -Path $dllPath

            $tagFile = [TagLib.File]::Create($filePath)

            if ($tagFile) {
                $tags = $tagFile.Tag
                if ($null -eq $tags) {
                    $tagFile.Tag = $tagFile.GetTag([TagLib.TagTypes]::Id3v2, $true)
                    $tags = $tagFile.Tag
                }

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
        Write-Host "Error writing metadata: $($_.Exception.Message)" -ForegroundColor Red
    }

    return $false
}

function Update-CommentsDisplay {
    $selectedCount = $controls.ListView.SelectedItems.Count

    if ($selectedCount -eq 1) {
        $index = $controls.ListView.SelectedItems[0].Index

        if ($index -lt $global:filteredTable.Count) {
            $file = $global:filteredTable[$index]

            if ($file -and $file.Path -and (Test-Path $file.Path)) {
                $extension = [System.IO.Path]::GetExtension($file.Path).ToLower()

                if ($extension -match "\.(m4a|mp3|ogg)$") {
                    $global:currentSelectedFile = $file
                    
                    if (-not $file.CommentsLoaded) {
                        $metadata = Read-FileMetadata $file.Path
                        $file.Comments = $metadata.Comments
                        $file.Duration = $metadata.Duration
                        $file.CommentsLoaded = $true

                        Update-ListViewTextColors
                    }
                    
                    $comments = $file.Comments
                    $global:originalCommentsText = $comments
                    $controls.CommentsBox.Text = $comments
                    $controls.SaveCommentsBtn.Enabled = $false
                } else {
                    $global:currentSelectedFile = $null
                    $global:originalCommentsText = ""
                    $controls.CommentsBox.Text = ""
                    $controls.SaveCommentsBtn.Enabled = $false
                }
            } else {
                $global:currentSelectedFile = $null
                $global:originalCommentsText = ""
                $controls.CommentsBox.Text = ""
                $controls.SaveCommentsBtn.Enabled = $false
            }
        } else {
            $global:currentSelectedFile = $null
            $global:originalCommentsText = ""
            $controls.CommentsBox.Text = ""
            $controls.SaveCommentsBtn.Enabled = $false
        }
    } else {
        $global:currentSelectedFile = $null
        $global:originalCommentsText = ""
        $controls.CommentsBox.Text = ""
        $controls.SaveCommentsBtn.Enabled = $false
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
    # Stop any ongoing background comment loading
    Stop-BackgroundCommentLoading
    
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
                Duration = 0
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
        
        # Apply year filter
        Apply-YearFilter

        Update-ListView
    } else {
        $global:fileTable = @()
        $global:filteredTable = @()
        $controls.SortCreatedRadio.Checked = $true
        Update-ListView
    }
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
            $global:filteredTable = $global:filteredTable | Sort-Object Duration -Descending
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
                $success = Write-FileMetadata $global:currentSelectedFile.Path $newComments
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

    $controls.CommentsBox.Add_TextChanged({
        if ($global:currentSelectedFile) {
            $currentText = $controls.CommentsBox.Text
            $originalText = $global:originalCommentsText

            $isDifferent = $currentText -ne $originalText
            $isNotEmpty = -not [string]::IsNullOrWhiteSpace($currentText)

            $controls.SaveCommentsBtn.Enabled = $isDifferent -and $isNotEmpty
        }
    })

    $controls.AboutLink.Add_Click({
        Start-Process "https://github.com/SVDotsenko/cleaner/blob/after-youtube/readme.md"
    })

    # Add year filter change handlers
    $controls.ThisYearRadio.Add_CheckedChanged({
        if ($controls.ThisYearRadio.Checked) {
            # Stop any ongoing background comment loading
            Stop-BackgroundCommentLoading

            # Apply filter and update main list
            Apply-YearFilter
            Update-ListView

            # Auto-load comments for visible items after year filter change
            if ($controls.ListView.Items.Count -gt 0) {
                Load-CommentsForVisibleItems
            }
        }
    })

    $controls.AllYearsRadio.Add_CheckedChanged({
        if ($controls.AllYearsRadio.Checked) {
            # Stop any ongoing background comment loading
            Stop-BackgroundCommentLoading

            # Apply filter and update main list
            Apply-YearFilter
            Update-ListView

            # Auto-load comments for visible items after year filter change
            if ($controls.ListView.Items.Count -gt 0) {
                Load-CommentsForVisibleItems
            }
        }
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
        
        $controls.ListView.AutoResizeColumn(0, [System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
        $controls.ListView.AutoResizeColumn(1, [System.Windows.Forms.ColumnHeaderAutoResizeStyle]::HeaderSize)
        $controls.ListView.Columns[2].Width = 100  # Fixed width for Created column
    }
})

$form.Topmost = $false

$form.Add_Shown({
    CreateControls
    Set-AllFonts $global:fontSize
    BindHandlers

    # Disable "All" radio button at startup
    $controls.AllYearsRadio.Enabled = $false

    Get-FilesFromFolder
    
    # Auto-load comments for visible items on startup
    Load-CommentsForVisibleItems

    # Ensure proper layout after form is fully shown
    LayoutOnlyFonts

    $form.Activate()
 })
 
 # Add form closing event to stop background loading
 $form.Add_FormClosing({
     Stop-BackgroundCommentLoading
 })
 
 [void]$form.ShowDialog()
