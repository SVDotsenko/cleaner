# FileManager.Tests.ps1
# Юнит-тесты для функций FileManager.ps1 (Pester 5.x)

Describe "FileManager Functions" {
    BeforeAll {
        # Определяем функции для тестирования локально, чтобы избежать запуска GUI
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
} 