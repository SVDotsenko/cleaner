Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

$global:fontSize = 14  # стартовый размер шрифта (можешь менять)
$global:fontFamily = "Segoe UI"

$form = New-Object Windows.Forms.Form
$form.Text = "File Manager"
$form.Width = 1200
$form.Height = 800
$form.MinimumSize = New-Object Drawing.Size(600,400)

# Контейнер для элементов
$controls = @{}

# Создаем ToolTip для подсказок
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 5000
$toolTip.InitialDelay = 1000
$toolTip.ReshowDelay = 500

# ====== Функция для создания и размещения элементов ======
function CreateControls {
    $gap = [int]($global:fontSize * 0.8)
    $btnH = [int]($global:fontSize * 2.2)
    $btnVPad = [int]($global:fontSize * 0.4)

    $y = $gap
    $x = $gap

    # Первая строка
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
#    $controls.DeleteToTrashCheckBox.Text = "To Trash"
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

    # Вторая строка
    $y2 = $y + $btnH + $gap
    $x2 = $gap

    $controls.DecreaseFontBtn = New-Object Windows.Forms.Button
    $controls.DecreaseFontBtn.Text = "A-"
    $form.Controls.Add($controls.DecreaseFontBtn)
    $controls.DecreaseFontBtn.SetBounds($x2, $y2, 40 + $global:fontSize, $btnH)
    $x2 += $controls.DecreaseFontBtn.Width + $gap

    $controls.IncreaseFontBtn = New-Object Windows.Forms.Button
    $controls.IncreaseFontBtn.Text = "A+"
    $form.Controls.Add($controls.IncreaseFontBtn)
    $controls.IncreaseFontBtn.SetBounds($x2, $y2, 40 + $global:fontSize, $btnH)
    $x2 += $controls.IncreaseFontBtn.Width + $gap

    $controls.TotalFilesLabel = New-Object Windows.Forms.Label
    $controls.TotalFilesLabel.Text = "Total files: 0"
    $controls.TotalFilesLabel.AutoSize = $true
    $form.Controls.Add($controls.TotalFilesLabel)
    # Сразу ставим шрифт, чтобы .Width был верным
    $font = New-Object System.Drawing.Font($global:fontFamily, $global:fontSize)
    $controls.TotalFilesLabel.Font = $font
    $form.PerformLayout()
    # Центрируем по вертикали относительно кнопок
    $controls.TotalFilesLabel.Top = $y2 + ($btnH - $controls.TotalFilesLabel.Height) / 2
    $controls.TotalFilesLabel.Left = $x2
    $form.PerformLayout()
    $x2 += $controls.TotalFilesLabel.PreferredWidth + $gap

    $controls.TotalSizeLabel = New-Object Windows.Forms.Label
    $controls.TotalSizeLabel.Text = "Total size: 0 MB"
    $controls.TotalSizeLabel.AutoSize = $true
    $form.Controls.Add($controls.TotalSizeLabel)
    $controls.TotalSizeLabel.Font = $font
    $form.PerformLayout()
    # Центрируем по вертикали относительно кнопок
    $controls.TotalSizeLabel.Top = $y2 + ($btnH - $controls.TotalSizeLabel.Height) / 2
    $controls.TotalSizeLabel.Left = $x2

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
    $controls.ListView.Columns.Add("File Name", -1) | Out-Null  # -1 для автоматического размера
    $controls.ListView.Columns.Add("MB", 100) | Out-Null
    $controls.ListView.Columns.Add("Created", 100) | Out-Null

    # Set tooltips for UI elements
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

    # В CreateControls добавим чекбокс
    $controls.ShowFullNameCheckBox = New-Object Windows.Forms.CheckBox
    $controls.ShowFullNameCheckBox.Text = "Show full name"
    $controls.ShowFullNameCheckBox.Checked = $false
    $controls.ShowFullNameCheckBox.AutoSize = $true
    $form.Controls.Add($controls.ShowFullNameCheckBox)
    $controls.ShowFullNameCheckBox.SetBounds($x, $y, 140 + $global:fontSize*2, $btnH)
    $x += $controls.ShowFullNameCheckBox.Width + $gap
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

    $controls.DeleteToTrashCheckBox.SetBounds($x, $y, 100 + $global:fontSize*2, $btnH)
    $x += $controls.DeleteToTrashCheckBox.Width + $gap

    $controls.SortNameBtn.SetBounds($x, $y, 110 + $global:fontSize*2, $btnH)
    $x += $controls.SortNameBtn.Width + $gap

    $controls.SortSizeBtn.SetBounds($x, $y, 110 + $global:fontSize*2, $btnH)
    $x += $controls.SortSizeBtn.Width + $gap

    $controls.SortCreatedBtn.SetBounds($x, $y, 110 + $global:fontSize*2, $btnH)
    $x += $controls.SortCreatedBtn.Width + $gap

    $y2 = $y + $btnH + $gap
    $x2 = $gap

    $controls.DecreaseFontBtn.SetBounds($x2, $y2, 40 + $global:fontSize, $btnH)
    $x2 += $controls.DecreaseFontBtn.Width + $gap

    $controls.IncreaseFontBtn.SetBounds($x2, $y2, 40 + $global:fontSize, $btnH)
    $x2 += $controls.IncreaseFontBtn.Width + $gap

    $controls.TotalFilesLabel.Font = $font
    $form.PerformLayout()
    # Центрируем по вертикали относительно кнопок
    $controls.TotalFilesLabel.Top = $y2 + ($btnH - $controls.TotalFilesLabel.Height) / 2
    $controls.TotalFilesLabel.Left = $x2
    $form.PerformLayout()
    $x2 += $controls.TotalFilesLabel.PreferredWidth + $gap

    $controls.TotalSizeLabel.Font = $font
    $form.PerformLayout()
    # Центрируем по вертикали относительно кнопок
    $controls.TotalSizeLabel.Top = $y2 + ($btnH - $controls.TotalSizeLabel.Height) / 2
    $controls.TotalSizeLabel.Left = $x2

    # Таблица
    $gap = [int]($global:fontSize * 0.8)
    $controls.ListView.Left = $gap
    $controls.ListView.Top = $y2 + $btnH + $gap
    $controls.ListView.Width = $form.ClientSize.Width - $gap*2
    $controls.ListView.Height = $form.ClientSize.Height - $controls.ListView.Top - $gap
    $controls.ListView.Font = $font
    $controls.SelectFolder.Font = $font
    $controls.DeleteBtn.Font = $font
    $controls.DeleteToTrashCheckBox.Font = $font
    $controls.SortNameBtn.Font = $font
    $controls.SortSizeBtn.Font = $font
    $controls.SortCreatedBtn.Font = $font
    $controls.TotalFilesLabel.Font = $font
    $controls.TotalSizeLabel.Font = $font
    $controls.DecreaseFontBtn.Font = $font
    $controls.IncreaseFontBtn.Font = $font
    # В LayoutOnlyFonts добавим шрифт для чекбокса
    $controls.ShowFullNameCheckBox.Font = $font
}

# ====== Служебные функции ======
function Set-AllFonts($fontSize) {
    $font = New-Object System.Drawing.Font($global:fontFamily, $fontSize)
    foreach ($ctrl in $controls.Values) { $ctrl.Font = $font }
    $form.Font = $font
}

function Refresh-InfoLabels {
    # Check if any files are selected
    $selectedCount = $controls.ListView.SelectedItems.Count
    
    if ($selectedCount -gt 0) {
        # Show statistics for selected files
        $sum = 0
        foreach ($selectedItem in $controls.ListView.SelectedItems) {
            $fileIndex = $selectedItem.Index
            $file = $global:filteredTable[$fileIndex]
            $sum += $file.SizeMB
        }
        $sum = [math]::Round($sum, 2)
        $controls.TotalFilesLabel.Text = "Selected files: $selectedCount"
        $controls.TotalSizeLabel.Text = "Selected size: $sum MB"
    } else {
        # Show statistics for all files (current behavior)
        $count = $global:filteredTable.Count
        $sum = 0
        if ($count -gt 0) {
            $sum = ($global:filteredTable | Measure-Object -Property SizeMB -Sum).Sum
            $sum = [math]::Round($sum, 2)
        }
        $controls.TotalFilesLabel.Text = "Total files: $count"
        $controls.TotalSizeLabel.Text = "Total size: $sum MB"
    }
}

function Refresh-ListView {
    $controls.ListView.Items.Clear()
    foreach ($file in $global:filteredTable) {
        $displayName = if ($controls.ShowFullNameCheckBox.Checked) { $file.OrigName } else { $file.Name }
        $item = New-Object Windows.Forms.ListViewItem($displayName)
        $item.SubItems.Add("$($file.SizeMB)")
        $displayDate = Format-ExtractedDate $file.DisplayDate
        $item.SubItems.Add($displayDate)
        $controls.ListView.Items.Add($item) | Out-Null
    }
    $controls.DeleteBtn.Enabled = $controls.ListView.Items.Count -gt 0 -and $controls.ListView.SelectedItems.Count -gt 0
    # Автоматически подстраиваем ширину столбца File Name под контент
    $controls.ListView.AutoResizeColumn(0, [System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
    Refresh-InfoLabels
}

# ====== Функция для извлечения даты из имени файла ======
function Get-DateFromFileName($fileName, $extension) {
    $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    
    switch ($extension.ToLower()) {
        ".m4a" {
            # Ищем последнюю группу из 6 цифр подряд (формат YYMMDD)
            if ($nameWithoutExt -match ".*?(\d{6})(?:_\d{6})?$") {
                $dateStr = $matches[1]
                $year = "20" + $dateStr.Substring(0, 2)
                $month = $dateStr.Substring(2, 2)
                $day = $dateStr.Substring(4, 2)
                try {
                    return [DateTime]::ParseExact("$year-$month-$day", "yyyy-MM-dd", $null)
                } catch {
                    return $null
                }
            }
        }
        ".ogg" {
            # Ищем паттерн YYYY_MM_DD
            if ($nameWithoutExt -match ".*?(\d{4}_\d{2}_\d{2})") {
                $dateStr = $matches[1]
                try {
                    return [DateTime]::ParseExact($dateStr, "yyyy_MM_dd", $null)
                } catch {
                    return $null
                }
            }
        }
        ".mp3" {
            # Ищем первые 8 цифр подряд, начинающиеся с "20" (формат YYYYMMDD)
            if ($nameWithoutExt -match "(20\d{6})") {
                $dateStr = $matches[1]
                $year = $dateStr.Substring(0, 4)
                $month = $dateStr.Substring(4, 2)
                $day = $dateStr.Substring(6, 2)
                try {
                    return [DateTime]::ParseExact("$year-$month-$day", "yyyy-MM-dd", $null)
                } catch {
                    return $null
                }
            }
        }
    }
    
    return $null
}

# ====== Функция для форматирования даты для отображения ======
function Format-ExtractedDate($date) {
    if ($date -ne $null) {
        return $date.ToString("dd.MM.yy")
    }
    return "N/A"
}

# ====== Функция для извлечения отображаемого имени файла ======
function Get-DisplayNameFromFileName($fileName) {
    $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    if ($nameWithoutExt.Length -gt 0 -and [char]::IsLetter($nameWithoutExt[0])) {
        # Если первый символ - буква, берём все начальные буквы подряд
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
        # Если первый символ не буква, возвращаем имя файла с расширением
        return $fileName
    }
}

$global:folderPath = "G:\My Drive\recordings"
$global:fileTable = @()
$global:filteredTable = @()
$global:activeSortButton = $null  # Отслеживает текущую активную кнопку сортировки

function Rename-CallRecordingFiles {
    if (-not (Test-Path $global:folderPath)) {
        return
    }
    
    $files = Get-ChildItem -Path $global:folderPath -File
    $renamedCount = 0
    
    foreach ($file in $files) {
        if ($file.Name.StartsWith("Call recording ")) {
            $newName = $file.Name.Substring(15)  # Remove "Call recording " (15 characters)
            if (-not [string]::IsNullOrWhiteSpace($newName)) {
                $newPath = Join-Path $file.Directory.FullName $newName
                try {
                    # Check if a file with the new name already exists
                    if (-not (Test-Path $newPath)) {
                        Rename-Item -Path $file.FullName -NewName $newName -Force
                        $renamedCount++
                    }
                } catch {
                    # Ignore errors for individual files and continue with others
                }
            }
        }
    }
    
    if ($renamedCount -gt 0) {
        [Windows.Forms.MessageBox]::Show("$renamedCount file(s) renamed (removed 'Call recording ' prefix).", "Rename Complete", 'OK', 'Information') | Out-Null
    }
}

function Load-FilesFromFolder {
    # First, rename any files that start with "Call recording "
    Rename-CallRecordingFiles

    $global:fileTable = @()
    if (Test-Path $global:folderPath) {
        $files = Get-ChildItem -Path $global:folderPath -File | Where-Object { $_.Extension -match "\.(m4a|mp3|ogg)$" }
        foreach ($file in $files) {
            # Извлекаем дату из имени файла
            $extractedDate = Get-DateFromFileName $file.Name $file.Extension
            # Если дата не извлечена, используем дату создания файла
            $displayDate = if ($extractedDate -ne $null) { $extractedDate } else { $file.CreationTime }
            # Извлекаем отображаемое имя
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
        # Устанавливаем кнопку Created как активную по умолчанию
        $global:activeSortButton = $controls.SortCreatedBtn
        # Применяем сортировку по дате по умолчанию
        $global:fileTable = $global:fileTable | Sort-Object DisplayDate -Descending
        Update-SortButtonStates
        Apply-Search
    } else {
        $global:fileTable = @()
        # Устанавливаем кнопку Created как активную по умолчанию
        $global:activeSortButton = $controls.SortCreatedBtn
        Update-SortButtonStates
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

function Move-FileToRecycleBin($filePath) {
    try {
        # Create Shell COM object
        $shell = New-Object -ComObject Shell.Application
        $item = $shell.Namespace(0).ParseName($filePath)
        $item.InvokeVerb("delete")
        return $true
    } catch {
        return $false
    }
}

function Update-SortButtonStates {
    # Сбрасываем все кнопки сортировки в обычное состояние
    $controls.SortNameBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
    $controls.SortNameBtn.BackColor = [System.Drawing.SystemColors]::Control
    
    $controls.SortSizeBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
    $controls.SortSizeBtn.BackColor = [System.Drawing.SystemColors]::Control
    
    $controls.SortCreatedBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
    $controls.SortCreatedBtn.BackColor = [System.Drawing.SystemColors]::Control
    
    # Устанавливаем эффект "нажатости" для активной кнопки
    if ($global:activeSortButton -ne $null) {
        $global:activeSortButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $global:activeSortButton.BackColor = [System.Drawing.Color]::LightBlue
        $global:activeSortButton.ForeColor = [System.Drawing.Color]::DarkBlue
    }
}

function BindHandlers {
    $controls.IncreaseFontBtn.Add_Click({
        if ($global:fontSize -lt 32) {
            $global:fontSize += 1
            Set-AllFonts $global:fontSize
            LayoutOnlyFonts
            $controls.ListView.AutoResizeColumn(0, [System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
        }
    })

    $controls.DecreaseFontBtn.Add_Click({
        if ($global:fontSize -gt 5) {
            $global:fontSize -= 1
            Set-AllFonts $global:fontSize
            LayoutOnlyFonts
            $controls.ListView.AutoResizeColumn(0, [System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
        }
    })

    $controls.SelectFolder.Add_Click({
        $dialog = New-Object Windows.Forms.FolderBrowserDialog
        $dialog.Description = "Select a folder"
        if ($dialog.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) {
            $global:folderPath = $dialog.SelectedPath
            Load-FilesFromFolder
        }
    })
    $controls.ListView.Add_SelectedIndexChanged({ 
        $controls.DeleteBtn.Enabled = $controls.ListView.SelectedItems.Count -gt 0
        Refresh-InfoLabels
    })
    $controls.SortNameBtn.Add_Click({ 
        $global:activeSortButton = $controls.SortNameBtn
        Update-SortButtonStates
        $global:filteredTable = $global:filteredTable | Sort-Object @{Expression="Name"; Ascending=$true}, @{Expression="DisplayDate"; Ascending=$false}
        Refresh-ListView 
    })
    $controls.SortSizeBtn.Add_Click({ 
        $global:activeSortButton = $controls.SortSizeBtn
        Update-SortButtonStates
        $global:filteredTable = $global:filteredTable | Sort-Object SizeMB -Descending
        Refresh-ListView 
    })
    $controls.SortCreatedBtn.Add_Click({ 
        $global:activeSortButton = $controls.SortCreatedBtn
        Update-SortButtonStates
        $global:filteredTable = $global:filteredTable | Sort-Object DisplayDate -Descending
        Refresh-ListView 
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
                    # Move to recycle bin
                    $success = Move-FileToRecycleBin $file.Path
                } else {
                    # Permanent deletion
                    Remove-Item -Path $file.Path -Force
                    $success = $true
                }
                
                if ($success) {
                    $deleted++
                    $global:fileTable = $global:fileTable | Where-Object { $_.Path -ne $file.Path }
                    $global:filteredTable = $global:filteredTable | Where-Object { $_.Path -ne $file.Path }
                }
            } catch {
                # ignore errors for now
            }
        }
        
        Refresh-ListView
        $actionText = if ($useTrash) { "moved to Recycle Bin" } else { "permanently deleted" }
        [Windows.Forms.MessageBox]::Show("$deleted file(s) $actionText.", "Done", 'OK', 'Information') | Out-Null
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

    # В BindHandlers добавим обработчик для чекбокса
    $controls.ShowFullNameCheckBox.Add_CheckedChanged({ Refresh-ListView })
}

$form.Add_Resize({
    if ($controls.ListView -ne $null) {
        $gap = [int]($global:fontSize * 0.8)
        $controls.ListView.Left = $gap
        $controls.ListView.Width = $form.ClientSize.Width - $gap*2
        $controls.ListView.Height = $form.ClientSize.Height - $controls.ListView.Top - $gap
        # Автоматически подстраиваем ширину столбца File Name под содержимое и доступное пространство
        $controls.ListView.AutoResizeColumn(0, [System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
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
