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

# Устанавливаем переменную для тестового режима
$env:FILEMANAGER_TEST_MODE = "true"

try {
    # Определяем корневую папку проекта
    $rootPath = Split-Path $PSScriptRoot -Parent
    $testsPath = $PSScriptRoot
    $sourceFile = Join-Path $rootPath "FileManager.ps1"

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

    # Анализируем результаты
    $coverage = $result.CodeCoverage
    $totalLines = $coverage.NumberOfCommandsAnalyzed
    $coveredLines = $coverage.NumberOfCommandsExecuted
    $coveragePercent = if ($totalLines -gt 0) {
        [math]::Round(($coveredLines / $totalLines) * 100, 2)
    } else { 0 }

    # Показываем результаты в консоли
    Write-Host "`n$("=" * 50)" -ForegroundColor Cyan
    Write-Host "📊 TEST RESULTS" -ForegroundColor Cyan
    Write-Host $("=" * 50) -ForegroundColor Cyan
    Write-Host "Tests: " -NoNewline -ForegroundColor White
    Write-Host "$($result.PassedCount) passed, $($result.FailedCount) failed" -ForegroundColor $(if ($result.FailedCount -eq 0) { 'Green' } else { 'Red' })
    Write-Host "Coverage: " -NoNewline -ForegroundColor White
    $coverageColor = if ($coveragePercent -ge 80) { 'Green' } elseif ($coveragePercent -ge 60) { 'Yellow' } else { 'Red' }
    Write-Host "$coveragePercent% ($coveredLines/$totalLines commands)" -ForegroundColor $coverageColor

    # Анализируем функции
    $sourceContent = Get-Content $sourceFile
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

    # Анализируем покрытие функций
    foreach ($func in $functions) {
        $funcMissedLines = $coverage.MissedCommands | Where-Object {
            $_.Line -ge $func.StartLine -and $_.Line -le $func.EndLine
        }
        $funcCoveredLines = $coverage.HitCommands | Where-Object {
            $_.Line -ge $func.StartLine -and $_.Line -le $func.EndLine
        }

        $func.MissedLines = if ($funcMissedLines) { $funcMissedLines.Count } else { 0 }
        $func.CoveredLines = if ($funcCoveredLines) { $funcCoveredLines.Count } else { 0 }
        $func.TotalLines = $func.MissedLines + $func.CoveredLines
    }

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
        .function-item { background-color: #f8f9fa; border: 1px solid #dee2e6; border-radius: 4px; padding: 15px; margin: 10px 0; }
        .function-name { color: #495057; font-weight: bold; font-size: 16px; margin-bottom: 8px; }
        .function-lines { color: #6c757d; font-size: 14px; margin-bottom: 5px; }
        .function-coverage { font-family: monospace; font-size: 12px; }
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
"@

    # Добавляем информацию о функциях
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
    Write-Host $("=" * 50) -ForegroundColor Cyan

    # Открываем отчёт в браузере
    if ($OpenReport) {
        Start-Process $htmlPath
        Write-Host "🌐 Opening report in browser..." -ForegroundColor Green
    }

} catch {
    Write-Host "❌ Error: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    # Очищаем переменную окружения
    Remove-Item env:FILEMANAGER_TEST_MODE -ErrorAction SilentlyContinue
}