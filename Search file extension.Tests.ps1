#Requires -Modules Pester, Toolbox.Remoting
#Requires -Version 5.1

BeforeAll {
    $realCmdLet = @{
        ImportExcel = Get-Command Import-Excel
    }

    $testInputFile = @{
        MaxConcurrentJobs = 1
        MailTo            = 'bob@contoso.com'
        AD                = @{
            OU = @('OU=Computer,DC=contoso,DC=com')
        }
        Path              = @{
            (New-Item 'TestDrive:\folder' -ItemType Directory).FullName = @('.txt')
        }
        ComputersNotInOu  = $null
    }

    $testOutParams = @{
        FilePath = (New-Item "TestDrive:\Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName  = 'Test (Brecht)'
        ImportFile  = $testOutParams.FilePath
        LogFolder   = New-Item 'TestDrive:\log' -ItemType Directory
        ScriptAdmin = 'mike@contoso.com'
    }

    $testLatestPSSessionConfiguration = Get-PSSessionConfiguration |
    Sort-Object -Property 'Name' -Descending |
    Select-Object -ExpandProperty 'Name' -First 1

    Mock Get-PowerShellConnectableEndpointNameHC {
        $testLatestPSSessionConfiguration
    }
    Mock Send-MailHC
    Mock Write-EventLog
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @('ScriptName', 'ImportFile') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'send an e-mail to the admin when' {
    BeforeAll {
        $MailAdminParams = {
            ($To -eq $testParams.ScriptAdmin) -and
            ($Priority -eq 'High') -and
            ($Subject -eq 'FAILURE')
        }
    }
    It 'the log folder cannot be created' {
        $testNewParams = Copy-ObjectHC $testParams
        $testNewParams.LogFolder = 'xxx::\\notExistingLocation'

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

        $testInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        . $testScript @testParams
    }
    It 'Invoke-Command is not called' {
        Should -Not -Invoke Invoke-Command
    }
}
Describe 'when computers are found in AD' {
    BeforeAll {
        Mock Get-ServersHC {
            @('PC1', 'PC2')
        }
        Mock Invoke-Command

        $testInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        . $testScript @testParams
    }
    Context 'call Get-PowerShellConnectableEndpointNameHC for each computer' {
        It '<_>' -ForEach @('PC1', 'PC2') {
            Should -Invoke Get-PowerShellConnectableEndpointNameHC -Times 1 -Exactly -Scope Describe -ParameterFilter {
                $ComputerName -eq $_
            }
        }
    }
    Context 'call Invoke-Command for each computer' {
        It '<_>' -ForEach @('PC1', 'PC2') {
            Should -Invoke Invoke-Command -Times 1 -Exactly -Scope Describe -ParameterFilter {
                ($ComputerName -eq $_) -and
                ($ArgumentList.Keys -eq $testInputFile.Path.Keys) -and
                ($ArgumentList.Values -eq '.txt') -and
                ($ConfigurationName -eq $testLatestPSSessionConfiguration)
            }
        }
    }
}
Describe 'when the script runs' {
    BeforeAll {
        Mock Get-ServersHC { $env:COMPUTERNAME }

        $testMatchingFiles = @(
            'TestDrive:\folder\1.txt',
            'TestDrive:\folder\2.txt',
            'TestDrive:\folder\3.txt'
        ) | ForEach-Object {
            (New-Item $_ -ItemType File).FullName
        }

        $testIgnoredFiles = @(
            'TestDrive:\folder\ignore.zip',
            'TestDrive:\folder\ignore.pst',
            'TestDrive:\ignore.txt'
        ) | ForEach-Object {
            (New-Item $_ -ItemType File).FullName
        }

        $testInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

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

            $actual = & $realCmdLet.ImportExcel -Path $testExcelLogFile.FullName -WorksheetName 'MatchingFiles'
        }
        It 'is saved in the log folder' {
            $testExcelLogFile | Should -Not -BeNullOrEmpty
        }
        It 'contains the files matching the extension' {
            $actual | Should -HaveCount $testMatchingFiles.Count

            $i = 0
            $testMatchingFiles | ForEach-Object {
                $actual[$i].ComputerName | Should -Be $env:COMPUTERNAME
                $actual[$i].Path | Should -Be $_
                $actual[$i].LastWriteTime | Should -Not -BeNullOrEmpty
                $actual[$i].'Size' | Should -Not -BeNullOrEmpty
                $i++
            }
        }
    }
    Context 'the exported Excel file worksheet PathExists' {
        BeforeAll {
            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx'

            $actual = & $realCmdLet.ImportExcel -Path $testExcelLogFile.FullName -WorksheetName 'PathExists'
        }
        It 'is saved in the log folder' {
            $testExcelLogFile | Should -Not -BeNullOrEmpty
        }
        It 'contains an overview of ComputerName, Path and Exists' {
            $actual.ComputerName | Should -Be $env:COMPUTERNAME
            $actual.Path | Should -Be $testInputFile.Path.Keys
            $actual.Exists | Should -BeTrue
        }
    }
    It 'a summary mail is sent to the user' {
        Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($To -eq 'bob@contoso.com') -and
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