<#
.SYNOPSIS
    Sequentially shuts down all Exchange servers in the organization from a jump host.

.DESCRIPTION
    This script is intended to run on a jump host that has the Exchange Management Tools installed.
    It loads the Exchange management cmdlets, discovers all Exchange servers in the organization,
    sorts them alphabetically, and shuts them down sequentially from the last server to the first.

    After each shutdown command is sent successfully, the script verifies that the target server
    no longer responds to ICMP ping. Because the script runs from a separate jump host, the final
    Exchange server can also be validated after shutdown.

.NOTES
    Requirements:
    - Run the script with administrative privileges.
    - The executing account must have Exchange read permissions and OS rights for remote shutdown.
    - Remote shutdown must be allowed by firewall / local policy on target servers.
    - ICMP echo must be enabled if ping validation is expected to succeed.

    Author: Steffen Meyer
    Cloud Solution Architect
    Microsoft Deutschland GmbH

    Version: 1.1.0
    Modified: 07.10.2026
 
    Version History:
    1.1.0 - 07.10.2026 - Added "-Force", $ScriptRootPath and coloured output
    1.0.0 - 25.05.2026 - Initial version
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
#[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$LogPath = "$PSScriptRoot\Logs\Exchange-SequentialShutdown_$(Get-Date -Format 'yyyyMMdd_HHmmss').log",

    [Parameter(Mandatory = $false)]
    [int]$PingRetryCount = 30,

    [Parameter(Mandatory = $false)]
    [int]$PingRetryIntervalSeconds = 10,

    [Parameter(Mandatory = $false)]
    [int]$RemoteShutdownGraceSeconds = 5,

    [Parameter(Mandatory = $false)]
    [switch]$ContinueOnFailure,

    [Parameter(Mandatory = $false)]
    [switch]$SkipPingCheck,

    [Parameter(Mandatory = $false)]
    [switch]$Force
)

begin {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    $logDirectory = Split-Path -Path $LogPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($logDirectory) -and -not (Test-Path -Path $logDirectory)) {
        New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
    }

    $script:ResultList = New-Object System.Collections.Generic.List[object]

    function Write-Log {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Message,

            [Parameter(Mandatory = $false)]
            [ValidateSet('INFO','WARN','ERROR','SUCCESS')]
            [string]$Level = 'INFO'
        )

        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $line = "[{0}] [{1}] {2}" -f $timestamp, $Level, $Message

        # coloured console output depending on the log level
        switch ($Level) {
            'WARN'    {Write-Host $line -ForegroundColor Yellow }
            'ERROR'   {Write-Host $line -ForegroundColor Red }
            'SUCCESS' {Write-Host $line -ForegroundColor Green }
            default   {Write-Host $line }
        }
        
        Add-Content -Path $LogPath -Value $line
    }

    function Add-Result {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Server,

            [Parameter(Mandatory = $true)]
            [string]$Phase,

            [Parameter(Mandatory = $true)]
            [string]$Status,

            [Parameter(Mandatory = $true)]
            [string]$Details
        )

        $script:ResultList.Add(
            [PSCustomObject]@{
                Timestamp = Get-Date
                Server    = $Server
                Phase     = $Phase
                Status    = $Status
                Details   = $Details
            }
        ) | Out-Null
    }

    function Import-ExchangeManagementTools {
        [CmdletBinding()]
        param()

        # If the cmdlet is already available, there is nothing else to do.
        if (Get-Command -Name Get-ExchangeServer -ErrorAction SilentlyContinue) {
            Write-Log -Message "Exchange cmdlets are already available in the current session." -Level SUCCESS
            return
        }

        # Try the Exchange snap-in first.
        $snapInNames = @(
            'Microsoft.Exchange.Management.PowerShell.SnapIn',
            'Microsoft.Exchange.Management.PowerShell.E2010',
            'Microsoft.Exchange.Management.PowerShell.E2013',
            'Microsoft.Exchange.Management.PowerShell.E2016',
            'Microsoft.Exchange.Management.PowerShell.E2019'
        )

        foreach ($snapInName in $snapInNames) {
            try {
                Add-PSSnapin -Name $snapInName -ErrorAction Stop
                if (Get-Command -Name Get-ExchangeServer -ErrorAction SilentlyContinue) {
                    Write-Log -Message "Loaded Exchange PowerShell snap-in: $snapInName" -Level SUCCESS
                    return
                }
            }
            catch {
                Write-Verbose "Snap-in $snapInName could not be loaded: $($_.Exception.Message)"
            }
        }

        # Try common Exchange module paths as a fallback.
        $candidateModulePaths = @(
            "$env:ProgramFiles\Microsoft\Exchange Server\V15\Bin\RemoteExchange.ps1",
            "$env:ProgramFiles\Microsoft\Exchange Server\V15\Bin\Exchange.ps1"
        )

        foreach ($modulePath in $candidateModulePaths) {
            try {
                if (Test-Path -Path $modulePath) {
                    . $modulePath
                    if (Get-Command -Name Get-ExchangeServer -ErrorAction SilentlyContinue) {
                        Write-Log -Message "Loaded Exchange management script: $modulePath" -Level SUCCESS
                        return
                    }
                }
            }
            catch {
                Write-Verbose "Exchange bootstrap script $modulePath could not be loaded: $($_.Exception.Message)"
            }
        }

        throw "Exchange Management Tools could not be loaded. 'Get-ExchangeServer' is not available in this session."
    }

    function Test-HostOffline {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$ComputerName,

            [Parameter(Mandatory = $true)]
            [int]$RetryCount,

            [Parameter(Mandatory = $true)]
            [int]$RetryIntervalSeconds
        )

        # The function returns $true once the host stops responding to ping.
        # If the host still responds after all retries, the function returns $false.
        for ($attempt = 1; $attempt -le $RetryCount; $attempt++) {
            try {
                $isAlive = Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -ErrorAction SilentlyContinue
            }
            catch {
                $isAlive = $false
            }

            if (-not $isAlive) {
                return $true
            }

            Start-Sleep -Seconds $RetryIntervalSeconds
        }

        return $false
    }

    function Invoke-RemoteShutdown {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$ComputerName,

            [Parameter(Mandatory = $true)]
            [int]$GraceSeconds
        )

        # shutdown.exe is typically more reliable for planned remote shutdown workflows.
        $arguments = "/m \\$ComputerName /s /f /t $GraceSeconds /d p:0:0 /c `"Planned power outage - Exchange sequential shutdown`""

        # To make script functional, the following "$process" section needs the be reactivated
        $process = Start-Process -FilePath "$env:SystemRoot\System32\shutdown.exe" `
            -ArgumentList $arguments `
            -NoNewWindow `
            -Wait `
            -PassThru

        return $process.ExitCode
    }

    function Stop-ExecutionIfRequired {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Message
        )

        if ($ContinueOnFailure) {
            Write-Log -Message $Message -Level WARN
        }
        else {
            throw $Message
        }
    }

    Write-Log -Message "==== Start Exchange sequential shutdown from jump host ====" -Level INFO
    Write-Log -Message "Jump host: $env:COMPUTERNAME" -Level INFO
    Write-Log -Message "Log file: $LogPath" -Level INFO
}

process {
    try {
        Write-Log -Message "Loading Exchange Management Tools..." -Level INFO
        Import-ExchangeManagementTools

        Write-Log -Message "Discovering Exchange servers by using Get-ExchangeServer..." -Level INFO
        $exchangeServers = @(Get-ExchangeServer | Sort-Object -Property Name)

        if (-not $exchangeServers -or $exchangeServers.Count -eq 0) {
            throw "No Exchange servers were found in the organization."
        }

        $serverNamesAscending  = @($exchangeServers | Select-Object -ExpandProperty Name)
        $serverNamesDescending = @($exchangeServers | Sort-Object -Property Name -Descending | Select-Object -ExpandProperty Name)

        Write-Log -Message ("Exchange servers (ascending): {0}" -f ($serverNamesAscending -join ', ')) -Level INFO
        Write-Log -Message ("Shutdown order (descending): {0}" -f ($serverNamesDescending -join ' -> ')) -Level INFO

        foreach ($server in $serverNamesDescending) {
            if (-not $PSCmdlet.ShouldProcess($server, "Initiate remote shutdown from jump host")) {
                Add-Result -Server $server -Phase 'ShutdownInitiate' -Status 'Skipped' -Details 'Skipped by WhatIf / ShouldProcess.'
                continue
            }

            Write-Log -Message "Initiating shutdown for server '$server'..." -Level INFO
            Add-Result -Server $server -Phase 'ShutdownInitiate' -Status 'Started' -Details 'Remote shutdown command is being initiated.'

            if ($Force)
            {
                try {
                    $exitCode = Invoke-RemoteShutdown -ComputerName $server -GraceSeconds $RemoteShutdownGraceSeconds
                }
                catch {
                    $message = "Failed to execute remote shutdown for '$server'. Error: $($_.Exception.Message)"
                    Add-Result -Server $server -Phase 'ShutdownInitiate' -Status 'Error' -Details $_.Exception.Message
                    Stop-ExecutionIfRequired -Message $message
                    continue
                }

                if ($exitCode -ne 0) {
                    $message = "Shutdown command for '$server' returned exit code $exitCode."
                    Write-Log -Message $message -Level ERROR
                    Add-Result -Server $server -Phase 'ShutdownInitiate' -Status 'Failed' -Details "shutdown.exe ExitCode=$exitCode"
                    Stop-ExecutionIfRequired -Message "Stopping because the shutdown command was not accepted for '$server'."
                    continue
                }
                Write-Log -Message "Shutdown command successfully sent to '$server'." -Level SUCCESS
                Add-Result -Server $server -Phase 'ShutdownInitiate' -Status 'Success' -Details "shutdown.exe ExitCode=$exitCode"

                if ($SkipPingCheck) {
                    Write-Log -Message "Ping validation was skipped for '$server'." -Level WARN
                    Add-Result -Server $server -Phase 'PingCheck' -Status 'Skipped' -Details 'Ping validation was skipped by parameter.'
                    continue
                }

                Write-Log -Message "Validating shutdown of '$server' by waiting for ICMP timeout..." -Level INFO
                $isOffline = Test-HostOffline -ComputerName $server `
                                              -RetryCount $PingRetryCount `
                                              -RetryIntervalSeconds $PingRetryIntervalSeconds

                if ($isOffline) {
                    Write-Log -Message "Server '$server' no longer responds to ping. Offline state confirmed." -Level SUCCESS
                    Add-Result -Server $server -Phase 'PingCheck' -Status 'Success' -Details 'Ping timeout confirmed.'
                }
                else {
                    $message = "Server '$server' still responds to ping after all retries."
                    Write-Log -Message $message -Level ERROR
                    Add-Result -Server $server -Phase 'PingCheck' -Status 'Failed' -Details 'Server still responds to ping.'
                    Stop-ExecutionIfRequired -Message "Stopping because offline validation failed for '$server'."
                }
            }
            else
            {
                Write-Log -Message "Real shutdowns are inactive, if the shutdown process is active, the script would shutdown ""$($server)"" here." -Level WARN
                Write-Log -Message "To perform actual shutdowns, rerun this script with the ""-Force"" parameter." -Level WARN
            }
        }
    }
    catch {
        Write-Log -Message $_.Exception.Message -Level ERROR
        throw
    }
}

end {
    if ($Force)
        {    
        try {
            Write-Log -Message "==== Summary ====" -Level INFO

            $summary = $script:ResultList | Sort-Object -Property Timestamp, Server, Phase
            $summary | Format-Table -AutoSize | Out-String | ForEach-Object {
                Add-Content -Path $LogPath -Value $_
            }

            $summary
        }
        catch {
            Write-Warning $_.Exception.Message
        }
    }
}