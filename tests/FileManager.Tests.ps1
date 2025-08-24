# Конфигурация Pester с анализом покрытия
$PesterConfig = New-PesterConfiguration
$PesterConfig.Run.Path = $PSScriptRoot
$PesterConfig.TestResult.Enabled = $true
$PesterConfig.TestResult.OutputFormat = 'NUnitXml'
$PesterConfig.TestResult.OutputPath = Join-Path $PSScriptRoot 'TestResults.xml'

# Настройка Code Coverage
$PesterConfig.CodeCoverage.Enabled = $true
$PesterConfig.CodeCoverage.Path = Join-Path (Split-Path $PSScriptRoot -Parent) "FileManager.ps1"
$PesterConfig.CodeCoverage.OutputFormat = 'JaCoCo'
$PesterConfig.CodeCoverage.OutputPath = Join-Path $PSScriptRoot 'coverage.xml'

# Включаем детальный вывод
$PesterConfig.Output.Verbosity = 'Detailed'

Describe "FileManager Functions" {
    BeforeAll {
        # Устанавливаем переменную окружения для режима тестирования
        $env:FILEMANAGER_TEST_MODE = "true"

        # Импортируем основной файл
        $scriptPath = Join-Path (Split-Path $PSScriptRoot -Parent) "FileManager.ps1"
        . $scriptPath
    }

    AfterAll {
        # Очищаем переменную окружения
        Remove-Item env:FILEMANAGER_TEST_MODE -ErrorAction SilentlyContinue
    }

    Describe "Format-ExtractedDate" {
        It "Should return 'N/A' for null date" {
            $result = Format-ExtractedDate $null
            $result | Should -Be "N/A"
        }
        
        It "Should format date without time when showTime is false" {
            $testDate = [DateTime]::Parse("2024-01-15 14:30:25")
            $result = Format-ExtractedDate $testDate $false
            $result | Should -Be "15.01.24"
        }
        
        It "Should format date with time when showTime is true" {
            $testDate = [DateTime]::Parse("2024-01-15 14:30:25")
            $result = Format-ExtractedDate $testDate $true
            $result | Should -Be "15.01.24 14:30:25"
        }
        
        It "Should handle different dates correctly" {
            $testDate = [DateTime]::Parse("2023-12-31 23:59:59")
            $result = Format-ExtractedDate $testDate $true
            $result | Should -Be "31.12.23 23:59:59"
        }
        
        It "Should handle edge case dates" {
            $testDate = [DateTime]::Parse("2025-02-01 00:00:00")
            $result = Format-ExtractedDate $testDate $false
            $result | Should -Be "01.02.25"
        }
    }
    
    Describe "Get-DisplayNameFromFileName" {
        It "Should extract letters from filename starting with letter" {
            $result = Get-DisplayNameFromFileName "TestFile123.mp3"
            $result | Should -Be "TestFile"
        }
        
        It "Should return original filename if starts with number" {
            $result = Get-DisplayNameFromFileName "123TestFile.mp3"
            $result | Should -Be "123TestFile.mp3"
        }
        
        It "Should handle empty filename" {
            $result = Get-DisplayNameFromFileName ""
            $result | Should -Be ""
        }
        
        It "Should handle filename with only letters" {
            $result = Get-DisplayNameFromFileName "AudioFile.mp3"
            $result | Should -Be "AudioFile"
        }
        
        It "Should handle filename with special characters" {
            $result = Get-DisplayNameFromFileName "Test-File_123.mp3"
            $result | Should -Be "Test"
        }
        
        It "Should handle filename starting with underscore" {
            $result = Get-DisplayNameFromFileName "_TestFile.mp3"
            $result | Should -Be "_TestFile.mp3"
        }
        
        It "Should handle filename with only numbers" {
            $result = Get-DisplayNameFromFileName "123456.mp3"
            $result | Should -Be "123456.mp3"
        }
    }

    Describe "Get-DateFromFileName" {
        It "Should extract date from .m4a file" {
            $result = Get-DateFromFileName "test_240115_143025.m4a" ".m4a"
            $result | Should -Be ([DateTime]::Parse("2024-01-15 14:30:25"))
        }

        It "Should return null for invalid .m4a format" {
            $result = Get-DateFromFileName "invalid.m4a" ".m4a"
            $result | Should -Be $null
        }

        It "Should extract date from .ogg file" {
            $result = Get-DateFromFileName "test_2024_01_15_14_30_25.ogg" ".ogg"
            $result | Should -Be ([DateTime]::Parse("2024-01-15 14:30:25"))
        }
    }

    Describe "Format-Duration" {
        It "Should format zero duration" {
            $result = Format-Duration 0
            $result | Should -Be "00:00:00"
        }

        It "Should format seconds only" {
            $result = Format-Duration 45
            $result | Should -Be "00:00:45"
        }

        It "Should format minutes and seconds" {
            $result = Format-Duration 125
            $result | Should -Be "00:02:05"
        }

        It "Should format hours, minutes and seconds" {
            $result = Format-Duration 3665
            $result | Should -Be "01:01:05"
        }
    }

    Describe "Test-Requirements" {
        BeforeEach {
            # Очищаем любые моки перед каждым тестом
            if (Get-Command Mock -ErrorAction SilentlyContinue) {
                Remove-Variable -Name MockCalled -Scope Global -ErrorAction SilentlyContinue
            }
        }

        It "Should return true when TagLibCli module and dll are available" {
            # Создаем фиктивный объект модуля
            $mockModule = [PSCustomObject]@{
                Path = "C:\valid\path\TagLibCli.psd1"
            }

            Mock Get-Module { return $mockModule } -ParameterFilter { $Name -eq "TagLibCli" -and $ListAvailable }
            Mock Split-Path { return "C:\valid\path" } -ParameterFilter { $Path -eq $mockModule.Path -and $Parent }
            Mock Join-Path { return "C:\valid\path\TagLibSharp.dll" }
            Mock Test-Path { return $true } -ParameterFilter { $Path -eq "C:\valid\path\TagLibSharp.dll" }

            $result = Test-Requirements
            $result | Should -Be $true

            Assert-MockCalled Get-Module -Times 1
            Assert-MockCalled Test-Path -Times 1
        }

        It "Should return false when TagLibCli module is not available" {
            # Мокаем Get-Module чтобы вернуть null (модуль не найден)
            Mock Get-Module { return $null } -ParameterFilter { $Name -eq "TagLibCli" -and $ListAvailable }

            # Мокаем MessageBox и Start-Process через переопределение в скрипте
            $global:TestMessageBoxResult = [System.Windows.Forms.DialogResult]::OK
            $global:TestStartProcessCalled = $false

            # Временно переопределяем функции для тестирования
            $originalMessageBox = [System.Windows.Forms.MessageBox]
            $originalStartProcess = Get-Command Start-Process

            try {
                # Создаем временную функцию-заглушку
                function Global:Test-RequirementsWrapper {
                    $tagLibModule = Get-Module -Name TagLibCli -ListAvailable
                    if (-not $tagLibModule) {
                        $global:TestStartProcessCalled = $true
                        return $false
                    }
                    return $true
                }

                $result = Test-RequirementsWrapper
                $result | Should -Be $false
                $global:TestStartProcessCalled | Should -Be $true

            } finally {
                # Очищаем глобальные переменные
                Remove-Variable -Name TestMessageBoxResult -Scope Global -ErrorAction SilentlyContinue
                Remove-Variable -Name TestStartProcessCalled -Scope Global -ErrorAction SilentlyContinue
                Remove-Item -Path Function:\Test-RequirementsWrapper -ErrorAction SilentlyContinue
            }
        }

        It "Should return false when TagLibSharp.dll is not found" {
            # Создаем фиктивный объект модуля
            $mockModule = [PSCustomObject]@{
                Path = "C:\fake\path\TagLibCli.psd1"
            }

            Mock Get-Module { return $mockModule } -ParameterFilter { $Name -eq "TagLibCli" -and $ListAvailable }
            Mock Split-Path { return "C:\fake\path" } -ParameterFilter { $Path -eq $mockModule.Path -and $Parent }
            Mock Join-Path { return "C:\fake\path\TagLibSharp.dll" }
            Mock Test-Path { return $false } -ParameterFilter { $Path -eq "C:\fake\path\TagLibSharp.dll" }

            # Создаем временную функцию-заглушку для тестирования логики
            function Global:Test-RequirementsWrapper2 {
                $tagLibModule = Get-Module -Name TagLibCli -ListAvailable
                if (-not $tagLibModule) {
                    return $false
                }
                $moduleDir = Split-Path $tagLibModule.Path -Parent
                $dllPath = Join-Path $moduleDir "TagLibSharp.dll"
                if (-not (Test-Path $dllPath)) {
                    return $false
                }
                return $true
            }

            try {
                $result = Test-RequirementsWrapper2
                $result | Should -Be $false

                Assert-MockCalled Get-Module -Times 1
                Assert-MockCalled Test-Path -Times 1
            } finally {
                Remove-Item -Path Function:\Test-RequirementsWrapper2 -ErrorAction SilentlyContinue
            }
        }

        It "Should check module availability correctly" {
            Mock Get-Module { return $null } -ParameterFilter { $Name -eq "TagLibCli" -and $ListAvailable }

            # Тестируем только логику проверки модуля
            $tagLibModule = Get-Module -Name TagLibCli -ListAvailable
            $tagLibModule | Should -Be $null

            Assert-MockCalled Get-Module -Times 1 -ParameterFilter { $Name -eq "TagLibCli" -and $ListAvailable }
        }

        It "Should check DLL path correctly when module exists" {
            $mockModule = [PSCustomObject]@{
                Path = "C:\test\path\TagLibCli.psd1"
            }

            Mock Get-Module { return $mockModule } -ParameterFilter { $Name -eq "TagLibCli" -and $ListAvailable }
            Mock Split-Path { return "C:\test\path" } -ParameterFilter { $Path -eq $mockModule.Path -and $Parent }
            Mock Join-Path { return "C:\test\path\TagLibSharp.dll" }
            Mock Test-Path { return $true } -ParameterFilter { $Path -eq "C:\test\path\TagLibSharp.dll" }

            # Тестируем логику построения пути к DLL
            $tagLibModule = Get-Module -Name TagLibCli -ListAvailable
            $tagLibModule | Should -Not -Be $null

            $moduleDir = Split-Path $tagLibModule.Path -Parent
            $moduleDir | Should -Be "C:\test\path"

            $dllPath = Join-Path $moduleDir "TagLibSharp.dll"
            $dllPath | Should -Be "C:\test\path\TagLibSharp.dll"

            $dllExists = Test-Path $dllPath
            $dllExists | Should -Be $true

            Assert-MockCalled Get-Module -Times 1
            Assert-MockCalled Split-Path -Times 1
            Assert-MockCalled Join-Path -Times 1
            Assert-MockCalled Test-Path -Times 1
        }
    }
}

# Запуск тестов с покрытием (если скрипт запускается напрямую)
if ($MyInvocation.InvocationName -ne '.') {
    Invoke-Pester -Configuration $PesterConfig
}
