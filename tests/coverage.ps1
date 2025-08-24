param(
    [switch]$OpenReport = $true
)

Write-Host "🧪 Starting tests with coverage analysis..." -ForegroundColor Cyan

# Проверяем и устанавливаем Pester
if (-not (Get-Module -Name Pester -ListAvailable)) {
    Write-Host "Installing Pester module..." -ForegroundColor Yellow
    Install-Module -Name Pester -Force -SkipPublisherCheck -Scope CurrentUser
}

Import-Module Pester -Force
$env:FILEMANAGER_TEST_MODE = "true"

try {
    $rootPath = Split-Path $PSScriptRoot -Parent
    $testsPath = $PSScriptRoot
    $sourceFile = Join-Path $rootPath "FileManager.ps1"
    $logFile = Join-Path $testsPath "log.txt"

    # Функция для записи в файл с временной меткой
    function Write-ToFile {
        param(
            [string]$Message,
            [string]$Color = "White"
        )
        $timestamp = Get-Date -Format "HH:mm:ss.fff"
        "[$timestamp] $Message" | Add-Content -Path $logFile -Encoding UTF8
    }

    # Очищаем файл log.txt и добавляем заголовок
    "=" * 80 | Set-Content -Path $logFile -Encoding UTF8
    Write-ToFile "🔍 PESTER COVERAGE DIAGNOSTIC OUTPUT"
    Write-ToFile "Started at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    "=" * 80 | Add-Content -Path $logFile -Encoding UTF8

    # Настраиваем конфигурацию Pester 5
    $configuration = New-PesterConfiguration
    $configuration.Run.Path = $testsPath
    $configuration.Run.PassThru = $true
    $configuration.CodeCoverage.Enabled = $true
    $configuration.CodeCoverage.Path = $sourceFile
    $configuration.CodeCoverage.OutputPath = Join-Path $testsPath "coverage.xml"
    $configuration.CodeCoverage.OutputFormat = 'JaCoCo'
    $configuration.Output.Verbosity = 'Detailed'

    # Запускаем тесты с покрытием
    Write-Host "Running tests..." -ForegroundColor Green
    $result = Invoke-Pester -Configuration $configuration

    # ДЕТАЛЬНАЯ ДИАГНОСТИКА ОБЪЕКТА ПОКРЫТИЯ - В ФАЙЛ
    Write-ToFile ""
    Write-ToFile "=" * 60
    Write-ToFile "🔍 DETAILED COVERAGE OBJECT ANALYSIS"
    Write-ToFile "=" * 60

    $coverage = $result.CodeCoverage

    Write-ToFile ""
    Write-ToFile "📋 Coverage Object Properties:"
    $coverage | Get-Member -MemberType Property | ForEach-Object {
        $propName = $_.Name
        $propValue = $coverage.$propName
        if ($propValue -is [Array]) {
            Write-ToFile "  $propName`: Array[$($propValue.Count)]"
        } else {
            Write-ToFile "  $propName`: $propValue"
        }
    }

    Write-ToFile ""
    Write-ToFile "📊 Basic Coverage Stats:"
    Write-ToFile "  NumberOfCommandsAnalyzed: $($coverage.NumberOfCommandsAnalyzed)"
    Write-ToFile "  NumberOfCommandsExecuted: $($coverage.NumberOfCommandsExecuted)"
    Write-ToFile "  NumberOfCommandsMissed: $($coverage.NumberOfCommandsMissed)"

    Write-ToFile ""
    Write-ToFile "🎯 HitCommands Analysis:"
    if ($coverage.CommandsExecuted) {
        Write-ToFile "  Count: $($coverage.CommandsExecuted.Count)"
        Write-ToFile "  Sample entries (first 5):"
        $coverage.CommandsExecuted | Select-Object -First 5 | ForEach-Object {
            Write-ToFile "    Line $($_.Line): $($_.Command)"
        }

        $hitLines = $coverage.CommandsExecuted | Sort-Object Line | Select-Object -ExpandProperty Line
        Write-ToFile "  Line numbers covered: $($hitLines -join ', ')"
    } else {
        Write-ToFile "  CommandsExecuted is null or empty!"
    }

    Write-ToFile ""
    Write-ToFile "❌ MissedCommands Analysis:"
    if ($coverage.CommandsMissed) {
        Write-ToFile "  Count: $($coverage.CommandsMissed.Count)"
        Write-ToFile "  Sample entries (first 5):"
        $coverage.CommandsMissed | Select-Object -First 5 | ForEach-Object {
            Write-ToFile "    Line $($_.Line): $($_.Command)"
        }

        $missedLines = $coverage.CommandsMissed | Sort-Object Line | Select-Object -ExpandProperty Line
        Write-ToFile "  Line numbers missed: $($missedLines -join ', ')"
    } else {
        Write-ToFile "  CommandsMissed is null or empty!"
    }

    Write-ToFile ""
    Write-ToFile "📁 Analyzed Files:"
    if ($coverage.AnalyzedFiles) {
        $coverage.AnalyzedFiles | ForEach-Object {
            Write-ToFile "  $($_.FullName)"
        }
    } else {
        Write-ToFile "  No analyzed files found!"
    }

    # ДОПОЛНИТЕЛЬНАЯ ДИАГНОСТИКА КО��АНД - В ФАЙЛ
    Write-ToFile ""
    Write-ToFile "🔍 DETAILED COMMANDS ANALYSIS:"
    Write-ToFile "All CommandsExecuted with file info:"
    if ($coverage.CommandsExecuted) {
        $coverage.CommandsExecuted | Select-Object -First 10 | ForEach-Object {
            Write-ToFile "  File: $($_.File)"
            Write-ToFile "  Line: $($_.Line) | Command: $($_.Command)"
            Write-ToFile "  ---"
        }
    }

    Write-ToFile ""
    Write-ToFile "All CommandsMissed with file info:"
    if ($coverage.CommandsMissed) {
        $coverage.CommandsMissed | Select-Object -First 10 | ForEach-Object {
            Write-ToFile "  File: $($_.File)"
            Write-ToFile "  Line: $($_.Line) | Command: $($_.Command)"
            Write-ToFile "  ---"
        }
    }

    # АНАЛИЗ ИСХОДНОГО ФАЙЛА - В ФАЙЛ
    Write-ToFile ""
    Write-ToFile "📝 SOURCE FILE ANALYSIS:"
    Write-ToFile "Checking if source file exists: $sourceFile"
    Write-ToFile "File exists: $(Test-Path $sourceFile)"

    if (Test-Path $sourceFile) {
        $sourceContent = Get-Content $sourceFile -Raw
        Write-ToFile "File size: $($sourceContent.Length) characters"

        # Показываем первые несколько строк функций
        Write-ToFile ""
        Write-ToFile "First few lines of each function:"
        $sourceLines = Get-Content $sourceFile
        for ($i = 0; $i -lt $sourceLines.Length; $i++) {
            if ($sourceLines[$i] -match '^\s*function\s+([a-zA-Z0-9_-]+)') {
                $funcName = $matches[1]
                Write-ToFile "  Function $funcName (line $($i + 1)):"

                # Показываем следующие 3 строки после объявления функции
                for ($j = 1; $j -le 3; $j++) {
                    if (($i + $j) -lt $sourceLines.Length) {
                        $nextLine = $sourceLines[$i + $j]
                        Write-ToFile "    $($i + $j + 1): $nextLine"
                    }
                }
                Write-ToFile ""
            }
        }
    }

    # ПРОВЕРКА КОНФИГУРАЦИИ PESTER - В ФАЙЛ
    Write-ToFile ""
    Write-ToFile "⚙️ PESTER CONFIGURATION CHECK:"
    Write-ToFile "CodeCoverage.Enabled: $($configuration.CodeCoverage.Enabled.Value)"
    Write-ToFile "CodeCoverage.Path: $($configuration.CodeCoverage.Path.Value)"
    Write-ToFile "CodeCoverage.OutputFormat: $($configuration.CodeCoverage.OutputFormat.Value)"
    Write-ToFile "Run.Path: $($configuration.Run.Path.Value)"

    # ПРОВЕРКА РЕЗУЛЬТАТА PESTER - В ФАЙЛ
    Write-ToFile ""
    Write-ToFile "📊 PESTER RESULT ANALYSIS:"
    Write-ToFile "Result object properties:"
    $result | Get-Member -MemberType Property | ForEach-Object {
        $propName = $_.Name
        Write-ToFile "  $propName"
    }

    Write-ToFile ""
    Write-ToFile "CodeCoverage object type: $($coverage.GetType().FullName)"
    Write-ToFile "CodeCoverage is null: $($coverage -eq $null)"

    # Анализируем результаты
    $totalLines = $coverage.CommandsAnalyzedCount
    $coveredLines = $coverage.CommandsExecutedCount
    $coveragePercent = if ($totalLines -gt 0) {
        [math]::Round(($coveredLines / $totalLines) * 100, 2)
    } else { 0 }

    # Показываем резу��ьтаты в консоли
    Write-Host "`n$("=" * 50)" -ForegroundColor Cyan
    Write-Host "📊 TEST RESULTS" -ForegroundColor Cyan
    Write-Host $("=" * 50) -ForegroundColor Cyan
    Write-Host "Tests: " -NoNewline -ForegroundColor White
    Write-Host "$($result.PassedCount) passed, $($result.FailedCount) failed" -ForegroundColor $(if ($result.FailedCount -eq 0) { 'Green' } else { 'Red' })
    Write-Host "Coverage: " -NoNewline -ForegroundColor White
    $coverageColor = if ($coveragePercent -ge 80) { 'Green' } elseif ($coveragePercent -ge 60) { 'Yellow' } else { 'Red' }
    Write-Host "$coveragePercent% ($coveredLines/$totalLines commands)" -ForegroundColor $coverageColor

    # Анализируем функции - В ФАЙЛ
    Write-ToFile ""
    Write-ToFile "=" * 60
    Write-ToFile "🔍 FUNCTION ANALYSIS DIAGNOSTIC"
    Write-ToFile "=" * 60

    $sourceContent = Get-Content $sourceFile
    Write-ToFile "Source file lines count: $($sourceContent.Length)"

    $functions = @()
    $currentFunction = $null

    for ($i = 0; $i -lt $sourceContent.Length; $i++) {
        $line = $sourceContent[$i]
        $lineNumber = $i + 1

        if ($line -match '^\s*function\s+([a-zA-Z0-9_-]+)') {
            if ($currentFunction) {
                $currentFunction.EndLine = $lineNumber - 1
                $functions += $currentFunction
            }

            $currentFunction = @{
                Name = $matches[1]
                StartLine = $lineNumber
                EndLine = $lineNumber
                CoveredLines = 0
                MissedLines = 0
                TotalLines = 0
            }
        }

        if ($currentFunction -and $line -match '^\s*}\s*$') {
            $currentFunction.EndLine = $lineNumber
            $functions += $currentFunction
            $currentFunction = $null
        }
    }

    if ($currentFunction) {
        $currentFunction.EndLine = $sourceContent.Length
        $functions += $currentFunction
    }

    # Анализируем покрытие функций - В ФАЙЛ
    Write-ToFile ""
    Write-ToFile "🎯 Function Coverage Analysis:"
    foreach ($func in $functions) {
        Write-ToFile ""
        Write-ToFile "  Function: $($func.Name)"
        Write-ToFile "    Lines: $($func.StartLine) - $($func.EndLine)"

        $funcMissedLines = $coverage.CommandsMissed | Where-Object {
            $_.Line -ge $func.StartLine -and $_.Line -le $func.EndLine
        }
        $funcCoveredLines = $coverage.CommandsExecuted | Where-Object {
            $_.Line -ge $func.StartLine -and $_.Line -le $func.EndLine
        }

        Write-ToFile "    Hit commands in range: $($funcCoveredLines.Count)"
        Write-ToFile "    Missed commands in range: $($funcMissedLines.Count)"

        if ($funcCoveredLines) {
            Write-ToFile "    Covered lines: $($funcCoveredLines.Line -join ', ')"
        }
        if ($funcMissedLines) {
            Write-ToFile "    Missed lines: $($funcMissedLines.Line -join ', ')"
        }

        $func.MissedLines = if ($funcMissedLines) { $funcMissedLines.Count } else { 0 }
        $func.CoveredLines = if ($funcCoveredLines) { $funcCoveredLines.Count } else { 0 }
        $func.TotalLines = $func.MissedLines + $func.CoveredLines

        Write-ToFile "    Final stats: $($func.CoveredLines)/$($func.TotalLines) lines covered"
    }

    Write-ToFile ""
    Write-ToFile "=" * 60
    Write-ToFile "🔍 DIAGNOSTIC COMPLETED"
    Write-ToFile "Completed at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-ToFile "=" * 60

    # Создаём HTML отчёт
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $progressBarColor = if ($coveragePercent -ge 80) { 'green' } elseif ($coveragePercent -ge 60) { 'orange' } else { 'red' }

    $htmlReport = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Code Coverage Report</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background-color: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background-color: white; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); padding: 30px; }
        h1 { color: #333; border-bottom: 3px solid #007acc; padding-bottom: 10px; margin-bottom: 30px; }
        .summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; margin-bottom: 30px; }
        .metric { background-color: #f8f9fa; padding: 20px; border-radius: 6px; border-left: 4px solid #007acc; }
        .metric h3 { margin: 0 0 10px 0; color: #495057; font-size: 14px; text-transform: uppercase; letter-spacing: 1px; }
        .metric .value { font-size: 28px; font-weight: bold; margin: 5px 0; }
        .metric .details { color: #6c757d; font-size: 14px; }
        .progress-bar { width: 100%; height: 8px; background-color: #e9ecef; border-radius: 4px; overflow: hidden; margin: 10px 0; }
        .progress-fill { height: 100%; transition: width 0.3s ease; }
        .green { color: #28a745; } .orange { color: #fd7e14; } .red { color: #dc3545; }
        .bg-green { background-color: #28a745; } .bg-orange { background-color: #fd7e14; } .bg-red { background-color: #dc3545; }
        .functions { margin-top: 30px; }
        .functions h2 { color: #333; border-bottom: 2px solid #007acc; padding-bottom: 10px; }
        .functions-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 15px; margin-top: 20px; }
        .function-item { background-color: #f8f9fa; border: 1px solid #dee2e6; border-radius: 4px; padding: 15px; min-height: 100px; display: flex; flex-direction: column; justify-content: space-between; }
        .function-name { color: #495057; font-weight: bold; font-size: 14px; margin-bottom: 8px; line-height: 1.2; overflow-wrap: break-word; }
        .function-lines { color: #6c757d; font-size: 12px; margin-bottom: 8px; }
        .function-coverage { font-family: monospace; font-size: 11px; margin-top: auto; }
        .timestamp { color: #6c757d; font-size: 12px; text-align: right; margin-top: 20px; border-top: 1px solid #dee2e6; padding-top: 10px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 Code Coverage Report</h1>

        <div class="summary">
            <div class="metric">
                <h3>Line Coverage</h3>
                <div class="value $coverageColor">$coveragePercent%</div>
                <div class="progress-bar">
                    <div class="progress-fill bg-$progressBarColor" style="width: $coveragePercent%"></div>
                </div>
                <div class="details">$coveredLines of $totalLines commands covered</div>
            </div>

            <div class="metric">
                <h3>Test Results</h3>
                <div class="value green">$($result.PassedCount)</div>
                <div class="details">tests passed, $($result.FailedCount) failed</div>
            </div>

            <div class="metric">
                <h3>Functions</h3>
                <div class="value">$($functions.Count)</div>
                <div class="details">total functions analyzed</div>
            </div>
        </div>

        <div class="functions">
            <h2>🔍 Functions Coverage</h2>
            <div class="functions-grid">
"@

    # Добавляем информацию о функциях в сетке
    $sortedFunctions = $functions | Sort-Object {
        if ($_.TotalLines -eq 0) { 100 }
        else { ($_.CoveredLines / $_.TotalLines) * 100 }
    }

    foreach ($func in $sortedFunctions) {
        $funcCoveragePercent = if ($func.TotalLines -gt 0) {
            [math]::Round(($func.CoveredLines / $func.TotalLines) * 100, 1)
        } else { 100 }

        $funcCoverageClass = if ($funcCoveragePercent -ge 80) { "green" }
                            elseif ($funcCoveragePercent -ge 60) { "orange" }
                            else { "red" }

        $htmlReport += @"
                <div class="function-item">
                    <div class="function-name">function $($func.Name)()</div>
                    <div class="function-lines">Lines: $($func.StartLine) - $($func.EndLine)</div>
                    <div class="function-coverage">
                        <span class="$funcCoverageClass">Coverage: $funcCoveragePercent%</span>
                        ($($func.CoveredLines)/$($func.TotalLines) commands covered)
                    </div>
                </div>
"@
    }

    $htmlReport += @"
            </div>
        </div>

        <div class="timestamp">
            Report generated on $timestamp
        </div>
    </div>
</body>
</html>
"@

    # Сохраняем HTML отчёт
    $htmlPath = Join-Path $PSScriptRoot 'coverage-report.html'
    $htmlReport | Out-File -FilePath $htmlPath -Encoding UTF8

    Write-Host "📄 HTML report saved to: tests\coverage-report.html" -ForegroundColor Cyan
    Write-Host "📄 XML coverage saved to: tests\coverage.xml" -ForegroundColor Cyan
    Write-Host "📄 Diagnostic output saved to: tests\log.txt" -ForegroundColor Yellow
    Write-Host $("=" * 50) -ForegroundColor Cyan

    # Открываем отчёт в браузере
    if ($OpenReport) {
        Start-Process $htmlPath
        Write-Host "🌐 Opening report in browser..." -ForegroundColor Green
    }

} catch {
    Write-Host "❌ Error: $($_.Exception.Message)" -ForegroundColor Red
    # Записываем ошибку в файл
    if ($logFile) {
        Write-ToFile "❌ ERROR: $($_.Exception.Message)"
        Write-ToFile "Stack trace: $($_.Exception.StackTrace)"
    }
} finally {
    # Очищаем переменную окружения
    Remove-Item env:FILEMANAGER_TEST_MODE -ErrorAction SilentlyContinue
}