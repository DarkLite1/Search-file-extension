<# 
    .SYNOPSIS   
        Search for specific file extensions.

    .DESCRIPTION
        find all files with the requested extension on the servers in AD and 
        send a report to the user with LastWriteTime, Size, ...

    .PARAMETER Path
        Combination of local paths and the exceptions to search for. Can be a of
        type hash table or PSCustomObject

    .PARAMETER OU
        Organizational unit in the active directory where to look for servers.

    .EXAMPLE
        @{
            'E:/DEPARTMENTS' = @('.pst')
        }
    
        Search for all files with extension '.pst' in the folder 
        'E:/DEPARTMENTS'.
#>
                
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String[]]$OU,
    [Parameter(Mandatory)]
    [String[]]$MailTo,
    [Parameter(Mandatory)]
    [HashTable]$Path,
    [String]$ComputersNotInOU,
    [String]$LogFolder = $env:POWERSHELL_LOG_FOLDER,
    [String]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)

Begin {
    $scriptBlock = {
        Param (
            [Parameter(Mandatory)]
            [HashTable]$Paths
        )

        $result = [PSCustomObject]@{
            ComputerName = $env:COMPUTERNAME
            Request      = $Paths
            PathExist    = @{}
            File         = @()
            Error        = @()
        }
        
        foreach ($path in $Paths.GetEnumerator()) {
            Try {
                $result.PathExist[$path.Key] = $false

                if (Test-Path -LiteralPath $path.Key -PathType Container) {
                    $result.PathExist[$path.Key] = $true
                    foreach ($extension in $path.Value) {
                        $params = @{
                            Path    = '\\?\{0}' -f $path.Key
                            Recurse = $true
                            Filter  = '*{0}' -f $extension
                            Force   = $true
                        }
                        $result.File += Get-ChildItem @params
                    }
                }
            }
            Catch {
                $result.Error += $_
                $Error.RemoveAt(0)
            }
        }

        $result
    }

    Try {
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        Get-ScriptRuntimeHC -Start
        $Error.Clear()

        #region Logging
        try {
            $LogParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $LogFile = New-LogFileNameHC @LogParams
        }
        Catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion

        $mailParams = @{ }
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    Try {
        #region Get computer names for servers
        Try {
            $getParams = @{
                OU = $OU
            }
            if ($ComputersNotInOU) {
                $getParams.Path = $ComputersNotInOU
            }
            $serverComputerNames = Get-ServersHC @getParams
        }
        Catch {
            throw "Failed retrieving the servers: $_"
        }
        #endregion

        #region Get files from remote machines
        $jobs = foreach ($computerName in $serverComputerNames) {
            $invokeParams = @{
                ComputerName = $computerName
                ScriptBlock  = $scriptBlock
                ArgumentList = $Path
                asJob        = $true
            }
            Invoke-Command @invokeParams
        }

        $jobResults = if ($jobs) { $jobs | Wait-Job | Receive-Job }
        #endregion

        $excelParams = @{
            Path          = $LogFile + '- Log.xlsx'
            AutoSize      = $true
            WorksheetName = $null
            TableName     = $null
            FreezeTopRow  = $true
        }

        #region Export matching files to Excel log file
        $matchingFilesToExport = foreach (
            $job in 
            $jobResults | Where-Object { $_.File }
        ) {
            $job.File | Select-Object -Property @{
                name = 'ComputerName'; expression = { $job.ComputerName }
            },
            @{name = 'Path'; expression = { $_.FullName.TrimStart('\\?\') } },
            CreationTime, LastWriteTime,
            @{Name = 'Size'; Expression = { [MATH]::Round($_.Length / 1GB, 2) } },
            @{name = 'Size_'; expression = { $_.Length / 1GB } }
        }

        if ($matchingFilesToExport) {
            $M = "Export $($jobResults.File.Count) rows to Excel sheet 'MatchingFiles'"
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
            
            $excelParams.WorksheetName = 'MatchingFiles'
            $excelParams.TableName = 'MatchingFiles'
            
            $matchingFilesToExport | Export-Excel @excelParams -AutoNameRange -CellStyleSB {
                Param (
                    $WorkSheet,
                    $TotalRows,
                    $LastColumn
                )

                @($WorkSheet.Names['Size'].Style).ForEach( {
                        $_.NumberFormat.Format = '?\ \G\B'
                        $_.HorizontalAlignment = 'Center'
                    })

                @($WorkSheet.Names['Size_'].Style).ForEach( {
                        $_.NumberFormat.Format = '?\ \K\B'
                        $_.HorizontalAlignment = 'Center'
                    })
            }

            $mailParams.Attachments = $excelParams.Path
        }
        #endregion

        #region Export search errors
        $searchErrors = foreach (
            $job in 
            $jobResults | Where-Object { $_.Error }
        ) {
            $job.Error | Select-Object -Property @{
                name = 'ComputerName'; expression = { $job.ComputerName }
            },
            @{name = 'Error'; expression = { $_ } }
        }

        if ($searchErrors) {
            $excelParams.WorksheetName = 'SearchErrors'
            $excelParams.TableName = 'SearchErrors'
            
            $searchErrors | Export-Excel @excelParams
        }
        #endregion

        #region Export path exists
        $pathExists = foreach (
            $job in 
            $jobResults | Where-Object { $_.PathExist }
        ) {
            $job.PathExist.GetEnumerator() | Select-Object -Property @{
                name = 'ComputerName'; expression = { $job.ComputerName }
            },
            @{name = 'Path'; expression = { $_.Key } },
            @{name = 'Exists'; expression = { $_.Value } }
        }

        if ($pathExists) {
            $excelParams.WorksheetName = 'PathExists'
            $excelParams.TableName = 'PathExists'
            
            $pathExists | Export-Excel @excelParams
        }
        #endregion

        #region Export general errors to Excel
        if ($Error.Exception.Message) {
            $Error.Exception.Message |
            Select-Object @{Name = 'Error'; Expression = { $_ } } |
            Export-Excel @excelParams -WorksheetName 'Error' -TableName 'Error'

            $mailParams.Attachments = $excelParams.Path
        }
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
    Finally {
        Get-Job | Remove-Job -Force
    }
}

end {
    Try {
        #region Send mail to user
        $searchFilters = ($Path.GetEnumerator() | ForEach-Object {
                "Path '{0}' extension '{1}'" -f $_.Key, $($_.Value -join "', '")
            }) -join '<br>'
                   
        $mailParams.Subject = "$($matchingFilesToExport.count) matching files"
        
        $errorMessage = $null
           
        if ($Error) {
            $mailParams.Priority = 'High'
            $mailParams.Subject = "$($Error.Count) errors, $($mailParams.Subject)"
            $errorMessage = "<p>Encountered <b>$($Error.Count) non terminating errors</b>. Check the 'Error' worksheet.</p>"
        }
        
        if ($folderRemovalErrors) {
            $mailParams.Priority = 'High'
            $mailParams.Subject += ", $($searchErrors.Count) search errors"
        }
           
        $table = "
        <table>
            <tr>
                <th>Servers scanned</th>
                <td>$($serverComputerNames.Count)</td>
            </tr>
            <tr>
                <th>Search filters</th>
                <td>$($searchFilters)</td>
            </tr>
            <tr>
                <th>Matching files found</th>
                <td>$($matchingFilesToExport.Count)</td>
            </tr>
            <tr>
                <th>Search errors</th>
                <td>$($searchErrors.Count)</td>
            </tr>
        </table>
        "
           
        $mailParams += @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Message   = "<p>Scan summary:</p>
                        $table
                        $errorMessage
                        <p><i>* Check the attachment for details</i></p>"
            LogFolder = $LogParams.LogFolder
            Header    = $ScriptName
            Save      = $LogFile + ' - Mail.html'
        }
           
        Get-ScriptRuntimeHC -Stop
        Send-MailHC @mailParams
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}