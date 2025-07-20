Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

# --- Глобальные параметры шрифта ---
$global:fontSize = 12
$global:fontFamily = "Segoe UI"

$form = New-Object Windows.Forms.Form
$form.Text = "File Manager"
$form.Width = 900
$form.Height = 610
$form.MinimumSize = New-Object Drawing.Size(500,370)

# --- Первая строка кнопок ---
$selectFolderBtn = New-Object Windows.Forms.Button
$selectFolderBtn.Text = "Select Folder"
$selectFolderBtn.Left = 10
$selectFolderBtn.Top = 10
# Ширина будет установлена в Set-AllFonts
$form.Controls.Add($selectFolderBtn)

$deleteBtn = New-Object Windows.Forms.Button
$deleteBtn.Text = "Delete Selected"
$deleteBtn.Left = 120
$deleteBtn.Top = 10
# Ширина будет установлена в Set-AllFonts
$deleteBtn.Enabled = $false
$form.Controls.Add($deleteBtn)

$sortNameBtn = New-Object Windows.Forms.Button
$sortNameBtn.Text = "Sort by Name"
$sortNameBtn.Left = 250
$sortNameBtn.Top = 10
# Ширина будет установлена в Set-AllFonts
$form.Controls.Add($sortNameBtn)

$sortSizeBtn = New-Object Windows.Forms.Button
$sortSizeBtn.Text = "Sort by Size"
$sortSizeBtn.Left = 370
$sortSizeBtn.Top = 10
# Ширина будет установлена в Set-AllFonts
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

# --- Вторая строка: поиск и кнопки шрифта ---
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
# Ширина будет установлена в Set-AllFonts
$form.Controls.Add($searchBtn)

$decreaseFontBtn = New-Object Windows.Forms.Button
$decreaseFontBtn.Text = "A-"
$decreaseFontBtn.Left = 410
$decreaseFontBtn.Top = 37
# Ширина будет установлена в Set-AllFonts
$form.Controls.Add($decreaseFontBtn)

$increaseFontBtn = New-Object Windows.Forms.Button
$increaseFontBtn.Text = "A+"
$increaseFontBtn.Left = 455
$increaseFontBtn.Top = 37
# Ширина будет установлена в Set-AllFonts
$form.Controls.Add($increaseFontBtn)

# --- Таблица файлов ---
$listView = New-Object Windows.Forms.ListView
$listView.View = 'Details'
$listView.FullRowSelect = $true
$listView.MultiSelect = $true
$listView.Left = 10
$listView.Top = 70
$listView.Width = 860
$listView.Height = 470
$listView.Columns.Add("File Name", 470) | Out-Null
$listView.Columns.Add("Size (MB)", 100) | Out-Null
$listView.Scrollable = $true
$listView.GridLines = $true
$listView.Sorting = 'None'
$listView.Anchor = "Top,Bottom,Left,Right"
$form.Controls.Add($listView)

# --- Глобальные переменные для файлов ---
$global:folderPath = "G:\My Drive\recordings"
$global:fileTable = @()
$global:filteredTable = @()

# --- Функция для изменения размера кнопок в соответствии с текстом ---
function Resize-ButtonToFitText($button, $font) {
    # Создать объект Graphics для измерения текста
    $g = [System.Drawing.Graphics]::FromHwnd($button.Handle)
    $size = $g.MeasureString($button.Text, $font)
    $g.Dispose()
    # Задать ширину кнопки с запасом (например, +20 пикселей)
    $button.Width = [Math]::Ceiling($size.Width) + 20
    # Опционально изменить высоту, если шрифт большой
    $button.Height = [Math]::Ceiling($size.Height) + 8
}

# --- Функция для корректировки позиций элементов первой строки ---
function Adjust-FirstRowLayout() {
    # Установка позиций для кнопок в первой строке
    $deleteBtn.Left = $selectFolderBtn.Left + $selectFolderBtn.Width + 10
    $sortNameBtn.Left = $deleteBtn.Left + $deleteBtn.Width + 10
    $sortSizeBtn.Left = $sortNameBtn.Left + $sortNameBtn.Width + 10

    # Статистика начинается после кнопок
    $totalFilesLabel.Left = $sortSizeBtn.Left + $sortSizeBtn.Width + 20
    $totalSizeLabel.Left = $totalFilesLabel.Left + $totalFilesLabel.Width + 15
}

# --- Функция для корректировки позиций элементов второй строки ---
function Adjust-SecondRowLayout() {
    # Поле поиска и кнопка
    $searchBox.Left = $searchLabel.Left + $searchLabel.Width + 5
    $searchBtn.Left = $searchBox.Left + $searchBox.Width + 10

    # Кнопки изменения шрифта
    $decreaseFontBtn.Left = $searchBtn.Left + $searchBtn.Width + 10
    $increaseFontBtn.Left = $decreaseFontBtn.Left + $decreaseFontBtn.Width + 5
}

# --- Функция установки шрифта ко всем элементам ---
function Set-AllFonts($fontSize) {
    $font = New-Object System.Drawing.Font($global:fontFamily, $fontSize)
    $form.Font = $font
    $selectFolderBtn.Font = $font
    $deleteBtn.Font = $font
    $sortNameBtn.Font = $font
    $sortSizeBtn.Font = $font
    $increaseFontBtn.Font = $font
    $decreaseFontBtn.Font = $font
    $totalFilesLabel.Font = $font
    $totalSizeLabel.Font = $font
    $searchLabel.Font = $font
    $searchBox.Font = $font
    $searchBtn.Font = $font
    $listView.Font = $font

    # Автоматически подгонять размер кнопок под текст
    Resize-ButtonToFitText $selectFolderBtn $font
    Resize-ButtonToFitText $deleteBtn $font
    Resize-ButtonToFitText $sortNameBtn $font
    Resize-ButtonToFitText $sortSizeBtn $font
    Resize-ButtonToFitText $searchBtn $font
    Resize-ButtonToFitText $increaseFontBtn $font
    Resize-ButtonToFitText $decreaseFontBtn $font

    # Изменение положения кнопок и надписей
    Adjust-FirstRowLayout
    Adjust-SecondRowLayout
}

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

$form.Add_Shown({
    Set-AllFonts $global:fontSize
    Load-FilesFromFolder
    # Гарантируем, что элементы UI расположены правильно
    Adjust-FirstRowLayout
    Adjust-SecondRowLayout
})

$increaseFontBtn.Add_Click({
    if ($global:fontSize -lt 32) {
        $global:fontSize += 1
        Set-AllFonts $global:fontSize
        $listView.Columns[0].Width = $listView.ClientSize.Width - $listView.Columns[1].Width - 20
    }
})

$decreaseFontBtn.Add_Click({
    if ($global:fontSize -gt 5) {
        $global:fontSize -= 1
        Set-AllFonts $global:fontSize
        $listView.Columns[0].Width = $listView.ClientSize.Width - $listView.Columns[1].Width - 20
    }
})

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
    # При изменении размера окна, также пересчитываем позиции элементов
    Adjust-FirstRowLayout
    Adjust-SecondRowLayout
})

$form.Topmost = $false
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
