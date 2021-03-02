#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $commandImportExcel = Get-Command Import-Excel

    $MailAdminParams = {
        ($To -eq $ScriptAdmin) -and ($Priority -eq 'High') -and 
        ($Subject -eq 'FAILURE')
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName = 'Test (Brecht)'
        MailTo     = @('bob@contoso.com')
        OU         = 'OU=Computer,DC=contoso,DC=com'
        Path       = @{
            (New-Item 'TestDrive:/folder' -ItemType Directory).FullName = 
            @('.txt')
        }
        LogFolder  = New-Item 'TestDrive:/log' -ItemType Directory
    }

    Mock Send-MailHC
    Mock Write-EventLog
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @('OU', 'MailTo', 'ScriptName', 'Path') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory | 
        Should -BeTrue
    }
}
Describe 'send an e-mail to the admin when' {
    It 'the log folder cannot be created' {
        $testNewParams = $testParams.clone()
        $testNewParams.LogFolder = 'xxx:://notExistingLocation'

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and 
            ($Message -like '*Failed creating the log folder*')
        }
    }
}
Describe 'when no servers are found in AD' {
    BeforeAll {
        Mock Get-ServersHC
        Mock Invoke-Command

        . $testScript @testParams
    }
    It 'Invoke-Command is not called' {
        Should -Not -Invoke Invoke-Command 
    }
} 
Describe 'when servers are found in AD' {
    Context 'Invoke-Command' {
        BeforeAll {
            Mock Get-ServersHC {
                @('PC1', 'PC2')
            }
            Mock Invoke-Command

            . $testScript @testParams
        }
        It 'is called for each computer to search for the same paths and extensions' {
            Should -Invoke Invoke-Command -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($ComputerName -eq 'PC1') -and
                ($ArgumentList.Keys -eq $testParams.Path.Keys) -and
                ($ArgumentList.Values -eq '.txt')
            }
            Should -Invoke Invoke-Command -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($ComputerName -eq 'PC2') -and
                ($ArgumentList.Keys -eq $testParams.Path.Keys) -and
                ($ArgumentList.Values -eq '.txt')
            }
        } 
    } 
    Context 'and the script runs' {
        BeforeAll {
            Mock Get-ServersHC { $env:COMPUTERNAME }
            
            $testMatchingFiles = @(
                'TestDrive:/folder/1.txt', 
                'TestDrive:/folder/2.txt', 
                'TestDrive:/folder/3.txt'
            ) | ForEach-Object {
                '\\?\' + (New-Item $_ -ItemType File).FullName
            }

            $testIgnoredFiles = @(
                'TestDrive:/folder/ignore.zip', 
                'TestDrive:/folder/ignore.pst',
                'TestDrive:/ignore.txt'
            ) | ForEach-Object {
                '\\?\' + (New-Item $_ -ItemType File).FullName
            }

            . $testScript @testParams
        }
        It 'files with matching extension are collected' {
            $jobResults.File.FullName | 
            Should -HaveCount $testMatchingFiles.Count

            $testMatchingFiles | ForEach-Object {
                $jobResults.File.FullName | Should -Contain $_
            }
        }
        It 'not matching file extensions and paths are ignored' {
            $testIgnoredFiles | ForEach-Object {
                $jobResults.File.FullName | Should -Not -Contain $_
            }
        }
        Context 'the exported Excel file worksheet MatchingFiles' {
            BeforeAll {
                $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx'

                $actual = & $commandImportExcel -Path $testExcelLogFile.FullName -WorksheetName 'MatchingFiles'
            }
            It 'is saved in the log folder' {
                $testExcelLogFile | Should -Not -BeNullOrEmpty
            }
            It 'contains the files matching the extension' {
                $actual | Should -HaveCount $testMatchingFiles.Count

                $i = 0
                $testMatchingFiles | ForEach-Object {
                    $actual[$i].ComputerName | Should -Be $env:COMPUTERNAME
                    $actual[$i].Path | Should -Be $_.TrimStart('\\?\')
                    $actual[$i].LastWriteTime | Should -Not -BeNullOrEmpty
                    $actual[$i].'Size' | Should -Not -BeNullOrEmpty
                    $i++
                }
            } 
        }
        Context 'the exported Excel file worksheet PathExists' {
            BeforeAll {
                $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx'

                $actual = & $commandImportExcel -Path $testExcelLogFile.FullName -WorksheetName 'PathExists'
            }
            It 'is saved in the log folder' {
                $testExcelLogFile | Should -Not -BeNullOrEmpty
            }
            It 'contains an overview of ComputerName, Path and Exists' {
                $actual.ComputerName | Should -Be $env:COMPUTERNAME
                $actual.Path | Should -Be $testParams.Path.Keys
                $actual.Exists | Should -BeTrue
            } 
        }
        It 'a summary mail is sent to the user' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
                ($To -eq 'bob@contoso.com') -and
                ($Bcc -eq $ScriptAdmin) -and
                ($Priority -ne 'High') -and
                ($Subject -eq '3 matching files') -and
                ($Attachments -like '*log.xlsx') -and
                ($Message -like '*
                *Servers scanned*1*
                *Search filters*folder*>*.txt*
                *Matching files found*3*
                *Search errors*0*')
            }
        }
    }
}

<# 


Invoke-Pester 'T:\Prod\Search file extension\Search file extension.Tests.ps1' -Output Detailed #>