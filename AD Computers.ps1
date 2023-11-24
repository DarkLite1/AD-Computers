#Requires -Version 5.1
#Requires -Modules Toolbox.HTML, Toolbox.EventLog, Toolbox.ActiveDirectory
#Requires -Modules ImportExcel

<#
    .SYNOPSIS
        Report about all the computer names in AD with their OS.

    .DESCRIPTION
        Report about all the computer names in AD with their operating system, install date, ...

    .PARAMETER ImportFile
        A .json file containing the script arguments.

    .PARAMETER LogFolder
        Location for the log files.
    #>

[CmdletBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName = 'AD Computers',
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\AD Reports\AD Computers\$ScriptName",
    [String[]]$ScriptAdmin = @(
        $env:POWERSHELL_SCRIPT_ADMIN,
        $env:POWERSHELL_SCRIPT_ADMIN_BACKUP
    )
)

Begin {
    Try {
        $Error.Clear()
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        Get-ScriptRuntimeHC -Start

        #region Logging
        try {
            $logParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $logFile = New-LogFileNameHC @LogParams
        }
        Catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion

        #region Import input file
        $File = Get-Content $ImportFile -Raw -EA Stop | ConvertFrom-Json

        if (-not ($MailTo = $File.MailTo)) {
            throw "Input file '$ImportFile': No 'MailTo' addresses found."
        }

        if (-not ($OUs = $File.AD.OU)) {
            throw "Input file '$ImportFile': No 'AD.OU' found."
        }
        #endregion
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
        $computers = Get-ADComputerHC -OU $OUs -EA Stop

        $mailParams = @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Subject   = "$($computers.count) computers found"
            Message   = "<p><b>$(@($computers).count) computers</b> found:</p>"
            LogFolder = $logParams.LogFolder
            Header    = $ScriptName
            Save      = $logFile + ' - Mail.html'
        }

        if ($computers) {
            Remove-Item $logFile -Force -EA Ignore

            $excelParams = @{
                Path          = $logFile + '.xlsx'
                AutoSize      = $true
                BoldTopRow    = $true
                FreezeTopRow  = $true
                WorkSheetName = 'Computers'
                TableName     = 'Computers'
                ErrorAction   = 'Stop'
            }
            $computers | Export-Excel @excelParams

            $mailParams.Attachments = $excelParams.Path

            $mailParams.Message += $computers | Group-Object OS |
            Select-Object @{ 
                Name       = 'Operating system'
                Expression = { $_.Name } 
            },
            @{
                Name       = 'Total'
                Expression = { $_.Count } 
            } | 
            Sort-Object 'Operating system' | 
            ConvertTo-Html -As Table -Fragment

            $mailParams.Message += "<p><i>* Check the attachment for details</i></p>"
        }

        $mailParams.Message += $OUs | ConvertTo-OuNameHC -OU | 
        Sort-Object | ConvertTo-HtmlListHC -Header 'Organizational units:'

        Get-ScriptRuntimeHC -Stop
        Send-MailHC @MailParams
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