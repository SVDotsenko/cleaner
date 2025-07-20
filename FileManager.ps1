Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

$global:fontSize = 14  # стартовый размер шрифта (можешь менять)
$global:fontFamily = "Segoe UI"

$form = New-Object Windows.Forms.Form
$form.Text = "File Manager"
$form.Width = 900
$form.Height = 610
$form.MinimumSize = New-Object Drawing.Size(600,400)

# Контейнер для элементов
$controls = @{}

# ====== Функция для создания и размещения элементов ======
function CreateControls {
    $gap = [int]($global:fontSize * 0.8)
    $btnH = [int]($global:fontSize * 2.2)
    $btnVPad = [int]($global:fontSize * 0.4)

    $y = $gap
    $x = $gap

    # Первая строка
    $controls.SelectFolder = New-Object Windows.Forms.Button
    $controls.SelectFolder.Text = "Select Folder"
    $form.Controls.Add($controls.SelectFolder)
    $controls.SelectFolder.SetBounds($x, $y, 100 + $global:fontSize*2, $btnH)
    $x += $controls.SelectFolder.Width + $gap

    $controls.DeleteBtn = New-Object Windows.Forms.Button
    $controls.DeleteBtn.Text = "Delete Selected"
    $controls.DeleteBtn.Enabled = $false
    $form.Controls.Add($controls.DeleteBtn)
    $controls.DeleteBtn.SetBounds($x, $y, 120 + $global:fontSize*2, $btnH)
    $x += $controls.DeleteBtn.Width + $gap

    $controls.SortNameBtn = New-Object Windows.Forms.Button
    $controls.SortNameBtn.Text = "Sort by Name"
    $form.Controls.Add($controls.SortNameBtn)
    $controls.SortNameBtn.SetBounds($x, $y, 110 + $global:fontSize*2, $btnH)
    $x += $controls.SortNameBtn.Width + $gap

    $controls.SortSizeBtn = New-Object Windows.Forms.Button
    $controls.SortSizeBtn.Text = "Sort by Size"
    $form.Controls.Add($controls.SortSizeBtn)
    $controls.SortSizeBtn.SetBounds($x, $y, 110 + $global:fontSize*2, $btnH)
    $x += $controls.SortSizeBtn.Width + $gap

    $controls.TotalFilesLabel = New-Object Windows.Forms.Label
    $controls.TotalFilesLabel.Text = "Total files: 0"
    $controls.TotalFilesLabel.AutoSize = $true
    $form.Controls.Add($controls.TotalFilesLabel)
    $controls.TotalFilesLabel.Top = $y + [int]($btnH / 3)
    $controls.TotalFilesLabel.Left = $x
    $x += $controls.TotalFilesLabel.PreferredWidth + $gap

    $controls.TotalSizeLabel = New-Object Windows.Forms.Label
    $controls.TotalSizeLabel.Text = "Total size: 0 MB"
    $controls.TotalSizeLabel.AutoSize = $true
    $form.Controls.Add($controls.TotalSizeLabel)
    $controls.TotalSizeLabel.Top = $y + [int]($btnH / 3)
    $controls.TotalSizeLabel.Left = $x

    # Вторая строка
    $y2 = $y + $btnH + $gap
    $x2 = $gap

    $controls.SearchLabel = New-Object Windows.Forms.Label
    $controls.SearchLabel.Text = "Filter:"
    $controls.SearchLabel.AutoSize = $true
    $form.Controls.Add($controls.SearchLabel)
    $controls.SearchLabel.Top = $y2 + $btnVPad
    # Сразу ставим шрифт, чтобы .Width был верным
    $font = New-Object System.Drawing.Font($global:fontFamily, $global:fontSize)
    $controls.SearchLabel.Font = $font
    $form.PerformLayout()
    $controls.SearchLabel.Left = $x2
    $form.PerformLayout()
    $x2 += $controls.SearchLabel.Width + $gap

    $controls.SearchBox = New-Object Windows.Forms.TextBox
    $form.Controls.Add($controls.SearchBox)
    $controls.SearchBox.SetBounds($x2, $y2, 250 + $global:fontSize*5, $btnH)
    $x2 += $controls.SearchBox.Width + $gap

    $controls.SearchBtn = New-Object Windows.Forms.Button
    $controls.SearchBtn.Text = "Search"
    $form.Controls.Add($controls.SearchBtn)
    $controls.SearchBtn.SetBounds($x2, $y2, 80 + $global:fontSize*2, $btnH)
    $x2 += $controls.SearchBtn.Width + $gap

    $controls.DecreaseFontBtn = New-Object Windows.Forms.Button
    $controls.DecreaseFontBtn.Text = "A-"
    $form.Controls.Add($controls.DecreaseFontBtn)
    $controls.DecreaseFontBtn.SetBounds($x2, $y2, 40 + $global:fontSize, $btnH)
    $x2 += $controls.DecreaseFontBtn.Width + $gap

    $controls.IncreaseFontBtn = New-Object Windows.Forms.Button
    $controls.IncreaseFontBtn.Text = "A+"
    $form.Controls.Add($controls.IncreaseFontBtn)
    $controls.IncreaseFontBtn.SetBounds($x2, $y2, 40 + $global:fontSize, $btnH)

    # Таблица файлов
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
    $controls.ListView.Columns.Add("File Name", 470) | Out-Null
    $controls.ListView.Columns.Add("Size (MB)", 100) | Out-Null
}

# ====== Пересчитывает расположение элементов и обновляет только шрифт ======
function LayoutOnlyFonts {
    $gap = [int]($global:fontSize * 0.8)
    $btnH = [int]($global:fontSize * 2.2)
    $btnVPad = [int]($global:fontSize * 0.4)

    $y = $gap
    $x = $gap

    $controls.SelectFolder.SetBounds($x, $y, 100 + $global:fontSize*2, $btnH)
    $x += $controls.SelectFolder.Width + $gap

    $controls.DeleteBtn.SetBounds($x, $y, 120 + $global:fontSize*2, $btnH)
    $x += $controls.DeleteBtn.Width + $gap

    $controls.SortNameBtn.SetBounds($x, $y, 110 + $global:fontSize*2, $btnH)
    $x += $controls.SortNameBtn.Width + $gap

    $controls.SortSizeBtn.SetBounds($x, $y, 110 + $global:fontSize*2, $btnH)
    $x += $controls.SortSizeBtn.Width + $gap

    $controls.TotalFilesLabel.Top = $y + [int]($btnH / 3)
    $controls.TotalFilesLabel.Left = $x
    $x += $controls.TotalFilesLabel.PreferredWidth + $gap

    $controls.TotalSizeLabel.Top = $y + [int]($btnH / 3)
    $controls.TotalSizeLabel.Left = $x

    $y2 = $y + $btnH + $gap
    $x2 = $gap

    $font = New-Object System.Drawing.Font($global:fontFamily, $global:fontSize)
    $controls.SearchLabel.Font = $font
    $form.PerformLayout()
    $controls.SearchLabel.Left = $x2
    $controls.SearchLabel.Top = $y2 + $btnVPad
    $form.PerformLayout()
    $x2 += $controls.SearchLabel.Width + $gap

    $controls.SearchBox.SetBounds($x2, $y2, 250 + $global:fontSize*5, $btnH)
    $x2 += $controls.SearchBox.Width + $gap

    $controls.SearchBtn.SetBounds($x2, $y2, 80 + $global:fontSize*2, $btnH)
    $x2 += $controls.SearchBtn.Width + $gap

    $controls.DecreaseFontBtn.SetBounds($x2, $y2, 40 + $global:fontSize, $btnH)
    $x2 += $controls.DecreaseFontBtn.Width + $gap

    $controls.IncreaseFontBtn.SetBounds($x2, $y2, 40 + $global:fontSize, $btnH)

    # Таблица
    $gap = [int]($global:fontSize * 0.8)
    $controls.ListView.Left = $gap
    $controls.ListView.Top = $y2 + $btnH + $gap
    $controls.ListView.Width = $form.ClientSize.Width - $gap*2
    $controls.ListView.Height = $form.ClientSize.Height - $controls.ListView.Top - $gap
    $controls.ListView.Font = $font
    $controls.SelectFolder.Font = $font
    $controls.DeleteBtn.Font = $font
    $controls.SortNameBtn.Font = $font
    $controls.SortSizeBtn.Font = $font
    $controls.TotalFilesLabel.Font = $font
    $controls.TotalSizeLabel.Font = $font
    $controls.SearchBox.Font = $font
    $controls.SearchBtn.Font = $font
    $controls.DecreaseFontBtn.Font = $font
    $controls.IncreaseFontBtn.Font = $font
}

# ====== Служебные функции ======
function Set-AllFonts($fontSize) {
    $font = New-Object System.Drawing.Font($global:fontFamily, $fontSize)
    foreach ($ctrl in $controls.Values) { $ctrl.Font = $font }
    $form.Font = $font
}

function Refresh-InfoLabels {
    $count = $global:filteredTable.Count
    $sum = 0
    if ($count -gt 0) {
        $sum = ($global:filteredTable | Measure-Object -Property SizeMB -Sum).Sum
        $sum = [math]::Round($sum, 2)
    }
    $controls.TotalFilesLabel.Text = "Total files: $count"
    $controls.TotalSizeLabel.Left = $controls.TotalFilesLabel.Left + $controls.TotalFilesLabel.PreferredWidth + ([int]($global:fontSize*0.8))
    $controls.TotalSizeLabel.Text = "Total size: $sum MB"
}

function Refresh-ListView {
    $controls.ListView.Items.Clear()
    foreach ($file in $global:filteredTable) {
        $item = New-Object Windows.Forms.ListViewItem($file.Name)
        $item.SubItems.Add("$($file.SizeMB)")
        $controls.ListView.Items.Add($item) | Out-Null
    }
    $controls.DeleteBtn.Enabled = $controls.ListView.Items.Count -gt 0 -and $controls.ListView.SelectedItems.Count -gt 0
    Refresh-InfoLabels
}

$global:folderPath = "G:\My Drive\recordings"
$global:fileTable = @()
$global:filteredTable = @()

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
    Refresh-ListView
}

function BindHandlers {
    $controls.IncreaseFontBtn.Add_Click({
        if ($global:fontSize -lt 32) {
            $global:fontSize += 1
            Set-AllFonts $global:fontSize
            LayoutOnlyFonts
            $controls.ListView.Columns[0].Width = $controls.ListView.ClientSize.Width - $controls.ListView.Columns[1].Width - 20
        }
    })

    $controls.DecreaseFontBtn.Add_Click({
        if ($global:fontSize -gt 5) {
            $global:fontSize -= 1
            Set-AllFonts $global:fontSize
            LayoutOnlyFonts
            $controls.ListView.Columns[0].Width = $controls.ListView.ClientSize.Width - $controls.ListView.Columns[1].Width - 20
        }
    })

    $controls.SearchBtn.Add_Click({ Apply-Search })
    $controls.SearchBox.Add_KeyDown({ param($sender, $e) if ($e.KeyCode -eq 'Return') { Apply-Search } })
    $controls.SelectFolder.Add_Click({
        $dialog = New-Object Windows.Forms.FolderBrowserDialog
        $dialog.Description = "Select a folder"
        if ($dialog.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) {
            $global:folderPath = $dialog.SelectedPath
            Load-FilesFromFolder
        }
    })
    $controls.ListView.Add_SelectedIndexChanged({ $controls.DeleteBtn.Enabled = $controls.ListView.SelectedItems.Count -gt 0 })
    $controls.SortNameBtn.Add_Click({ $global:filteredTable = $global:filteredTable | Sort-Object Name; Refresh-ListView })
    $controls.SortSizeBtn.Add_Click({ $global:filteredTable = $global:filteredTable | Sort-Object SizeMB -Descending; Refresh-ListView })

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

    $controls.ListView.Add_DoubleClick({
        if ($controls.ListView.SelectedItems.Count -eq 1) {
            $index = $controls.ListView.SelectedItems[0].Index
            $file = $global:filteredTable[$index]
            try {
                [System.Diagnostics.Process]::Start($file.Path) | Out-Null
            } catch {
                [Windows.Forms.MessageBox]::Show("Cannot open file: $($file.Path)", "Error", 'OK', 'Error') | Out-Null
            }
        }
    })
}

$form.Add_Resize({
    if ($controls.ListView -ne $null) {
        $gap = [int]($global:fontSize * 0.8)
        $controls.ListView.Left = $gap
        $controls.ListView.Width = $form.ClientSize.Width - $gap*2
        $controls.ListView.Height = $form.ClientSize.Height - $controls.ListView.Top - $gap
        $controls.ListView.Columns[0].Width = $controls.ListView.ClientSize.Width - $controls.ListView.Columns[1].Width - 20
    }
})

$form.Topmost = $false

$form.Add_Shown({
    CreateControls
    Set-AllFonts $global:fontSize
    BindHandlers
    Load-FilesFromFolder
    $form.Activate()
})

[void]$form.ShowDialog()
