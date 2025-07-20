Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

$form = New-Object Windows.Forms.Form
$form.Text = "File Manager"
$form.Width = 770
$form.Height = 580
$form.MinimumSize = New-Object Drawing.Size(500,370)

# ---- Кнопки в первой строке ----
$selectFolderBtn = New-Object Windows.Forms.Button
$selectFolderBtn.Text = "Select Folder"
$selectFolderBtn.Left = 10
$selectFolderBtn.Top = 10
$selectFolderBtn.Width = 100
$form.Controls.Add($selectFolderBtn)

$deleteBtn = New-Object Windows.Forms.Button
$deleteBtn.Text = "Delete Selected"
$deleteBtn.Left = 120
$deleteBtn.Top = 10
$deleteBtn.Width = 120
$deleteBtn.Enabled = $false
$form.Controls.Add($deleteBtn)

$sortNameBtn = New-Object Windows.Forms.Button
$sortNameBtn.Text = "Sort by Name"
$sortNameBtn.Left = 250
$sortNameBtn.Top = 10
$sortNameBtn.Width = 110
$form.Controls.Add($sortNameBtn)

$sortSizeBtn = New-Object Windows.Forms.Button
$sortSizeBtn.Text = "Sort by Size"
$sortSizeBtn.Left = 370
$sortSizeBtn.Top = 10
$sortSizeBtn.Width = 110
$form.Controls.Add($sortSizeBtn)

$totalFilesLabel = New-Object Windows.Forms.Label
$totalFilesLabel.Text = "Total files: 0"
$totalFilesLabel.Left = 500
$totalFilesLabel.Top = 15
$form.Controls.Add($totalFilesLabel)
$totalFilesLabel.AutoSize = $true

$totalSizeLabel = New-Object Windows.Forms.Label
$totalSizeLabel.Text = "Total size: 0 MB"
$totalSizeLabel.Left = $totalFilesLabel.Left + $totalFilesLabel.Width + 15
$totalSizeLabel.Top = 15
$form.Controls.Add($totalSizeLabel)
$totalSizeLabel.AutoSize = $true

# ---- Вторая строка: поиск ----
$searchLabel = New-Object Windows.Forms.Label
$searchLabel.Text = "Filter:"
$searchLabel.Left = 10
$searchLabel.Top = 42
$form.Controls.Add($searchLabel)
$searchLabel.AutoSize = $true

$searchBox = New-Object Windows.Forms.TextBox
$searchBox.Left = 60
$searchBox.Top = 38
$searchBox.Width = 250
$form.Controls.Add($searchBox)

$searchBtn = New-Object Windows.Forms.Button
$searchBtn.Text = "Search"
$searchBtn.Left = 320
$searchBtn.Top = 37
$searchBtn.Width = 80
$form.Controls.Add($searchBtn)

# ---- Таблица ----
$listView = New-Object Windows.Forms.ListView
$listView.View = 'Details'
$listView.FullRowSelect = $true
$listView.MultiSelect = $true
$listView.Left = 10
$listView.Top = 70
$listView.Width = 720
$listView.Height = 440
$listView.Columns.Add("File Name", 420) | Out-Null
$listView.Columns.Add("Size (MB)", 100) | Out-Null
$listView.Scrollable = $true
$listView.GridLines = $true
$listView.Sorting = 'None'
$listView.Anchor = "Top,Bottom,Left,Right"
$form.Controls.Add($listView)

$global:folderPath = "G:\My Drive\recordings"
$global:fileTable = @()
$global:filteredTable = @()

function Refresh-InfoLabels {
    $count = $global:filteredTable.Count
    $sum = 0
    if ($count -gt 0) {
        $sum = ($global:filteredTable | Measure-Object -Property SizeMB -Sum).Sum
        $sum = [math]::Round($sum, 2)
    }
    $totalFilesLabel.Text = "Total files: $count"
    $totalFilesLabel.AutoSize = $true
    $totalSizeLabel.Left = $totalFilesLabel.Left + $totalFilesLabel.Width + 15
    $totalSizeLabel.Text = "Total size: $sum MB"
    $totalSizeLabel.AutoSize = $true
}

function Refresh-ListView {
    $listView.Items.Clear()
    foreach ($file in $global:filteredTable) {
        $item = New-Object Windows.Forms.ListViewItem($file.Name)
        $item.SubItems.Add("$($file.SizeMB)")
        $listView.Items.Add($item) | Out-Null
    }
    $deleteBtn.Enabled = $listView.Items.Count -gt 0 -and $listView.SelectedItems.Count -gt 0
    Refresh-InfoLabels
}

function Load-FilesFromFolder {
    $global:fileTable = @()
    if (Test-Path $global:folderPath) {
        $files = Get-ChildItem -Path $global:folderPath -File
        foreach ($file in $files) {
            $global:fileTable += [PSCustomObject]@{
                Name   = $file.Name
                SizeMB = [math]::Round($file.Length / 1MB, 2)
                Path   = $file.FullName
            }
        }
        $global:fileTable = $global:fileTable | Sort-Object SizeMB -Descending
        Apply-Search
    } else {
        $global:fileTable = @()
        Apply-Search
    }
}

function Apply-Search {
    $pattern = $searchBox.Text
    if ([string]::IsNullOrWhiteSpace($pattern)) {
        $global:filteredTable = $global:fileTable
    } else {
        try {
            $regex = New-Object System.Text.RegularExpressions.Regex($pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            $global:filteredTable = $global:fileTable | Where-Object { $regex.IsMatch($_.Name) }
        } catch {
            $p = $pattern.ToLower()
            $global:filteredTable = $global:fileTable | Where-Object { $_.Name.ToLower() -like "*$p*" }
        }
    }
    Refresh-ListView
}

$form.Add_Shown({ Load-FilesFromFolder })

$searchBtn.Add_Click({
    Apply-Search
})

$searchBox.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq 'Return') { Apply-Search }
})

$selectFolderBtn.Add_Click({
    $dialog = New-Object Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Select a folder"
    if ($dialog.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) {
        $global:folderPath = $dialog.SelectedPath
        Load-FilesFromFolder
    }
})

$listView.Add_SelectedIndexChanged({
    $deleteBtn.Enabled = $listView.SelectedItems.Count -gt 0
})

$sortNameBtn.Add_Click({
    $global:filteredTable = $global:filteredTable | Sort-Object Name
    Refresh-ListView
})

$sortSizeBtn.Add_Click({
    $global:filteredTable = $global:filteredTable | Sort-Object SizeMB -Descending
    Refresh-ListView
})

$deleteBtn.Add_Click({
    $toDeleteIndexes = @()
    foreach ($item in $listView.SelectedItems) {
        $toDeleteIndexes += $item.Index
    }
    $toDeleteIndexes = $toDeleteIndexes | Sort-Object -Descending
    $deleted = 0
    foreach ($i in $toDeleteIndexes) {
        $file = $global:filteredTable[$i]
        try {
            [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
                $file.Path,
                'OnlyErrorDialogs',
                'SendToRecycleBin'
            ) | Out-Null
            $deleted++
            $global:fileTable = $global:fileTable | Where-Object { $_.Path -ne $file.Path }
            $global:filteredTable = $global:filteredTable | Where-Object { $_.Path -ne $file.Path }
        } catch {
            # ignore errors for now
        }
    }
    Refresh-ListView
    [Windows.Forms.MessageBox]::Show("$deleted file(s) sent to Recycle Bin.", "Done", 'OK', 'Information') | Out-Null
})

$listView.Add_DoubleClick({
    if ($listView.SelectedItems.Count -eq 1) {
        $index = $listView.SelectedItems[0].Index
        $file = $global:filteredTable[$index]
        try {
            [System.Diagnostics.Process]::Start($file.Path) | Out-Null
        } catch {
            [Windows.Forms.MessageBox]::Show("Cannot open file: $($file.Path)", "Error", 'OK', 'Error') | Out-Null
        }
    }
})

$form.Add_Resize({
    $listView.Columns[0].Width = $listView.ClientSize.Width - $listView.Columns[1].Width - 20
})

$form.Topmost = $true
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
