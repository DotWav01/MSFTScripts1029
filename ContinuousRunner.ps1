# ContinuousRunner.ps1
# Runs the GroupQuery script continuously every hour

param(
    [Parameter(Mandatory=$true)]
    [string]$GroupQueryScriptPath,
    
    [Parameter(Mandatory=$false)]
    [int]$IntervalHours = 1,
    
    [Parameter(Mandatory=$false)]
    [string]$LogPath,
    
    [Parameter(Mandatory=$false)]
    [hashtable]$ScriptParameters = @{},
    
    [Parameter(Mandatory=$false)]
    [switch]$RunOnce
)

# Validate script exists
if (!(Test-Path $GroupQueryScriptPath)) {
    Write-Error "GroupQuery script not found: $GroupQueryScriptPath"
    exit 1
}

# Set default log path
if ([string]::IsNullOrEmpty($LogPath)) {
    $ScriptDirectory = Split-Path $MyInvocation.MyCommand.Path
    $LogsFolder = Join-Path $ScriptDirectory "ContinuousRunnerLogs"
    if (!(Test-Path $LogsFolder)) {
        New-Item -ItemType Directory -Path $LogsFolder -Force | Out-Null
    }
    $LogPath = Join-Path $LogsFolder "ContinuousRunner_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
}

# Function to write log entries
function Write-RunnerLog {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Write-Host $logEntry
    Add-Content -Path $LogPath -Value $logEntry -ErrorAction SilentlyContinue
}

# Function to run the GroupQuery script
function Invoke-GroupQueryScript {
    Write-RunnerLog "=== Starting GroupQuery execution ==="
    
    try {
        # Build parameter list for splatting
        $paramString = ""
        if ($ScriptParameters.Count -gt 0) {
            $paramArray = @()
            foreach ($key in $ScriptParameters.Keys) {
                $value = $ScriptParameters[$key]
                if ($value -is [switch] -or $value -eq $true) {
                    $paramArray += "-$key"
                } elseif ($value -eq $false) {
                    $paramArray += "-$key:`$false"
                } else {
                    $paramArray += "-$key `"$value`""
                }
            }
            $paramString = $paramArray -join " "
        }
        
        Write-RunnerLog "Executing: $GroupQueryScriptPath $paramString"
        
        # Execute the script
        if ($ScriptParameters.Count -gt 0) {
            & $GroupQueryScriptPath @ScriptParameters
        } else {
            & $GroupQueryScriptPath
        }
        
        if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq $null) {
            Write-RunnerLog "GroupQuery completed successfully" "SUCCESS"
        } else {
            Write-RunnerLog "GroupQuery completed with exit code: $LASTEXITCODE" "WARN"
        }
    }
    catch {
        Write-RunnerLog "Error executing GroupQuery: $($_.Exception.Message)" "ERROR"
    }
    
    Write-RunnerLog "=== GroupQuery execution completed ==="
}

# Main execution
Write-RunnerLog "=== CONTINUOUS RUNNER STARTED ==="
Write-RunnerLog "GroupQuery Script: $GroupQueryScriptPath"
Write-RunnerLog "Interval: $IntervalHours hour(s)"
Write-RunnerLog "Log file: $LogPath"
Write-RunnerLog "Run once mode: $RunOnce"

if ($ScriptParameters.Count -gt 0) {
    Write-RunnerLog "Script parameters:"
    foreach ($key in $ScriptParameters.Keys) {
        Write-RunnerLog "  $key = $($ScriptParameters[$key])"
    }
}

# Handle Ctrl+C gracefully
$cancelPressed = $false
Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
    $global:cancelPressed = $true
    Write-Host "`nShutdown requested - stopping after current execution..." -ForegroundColor Yellow
}

try {
    do {
        $startTime = Get-Date
        Write-RunnerLog "Current time: $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))"
        
        # Run the GroupQuery script
        Invoke-GroupQueryScript
        
        # Exit if running once
        if ($RunOnce) {
            Write-RunnerLog "Single execution completed, exiting"
            break
        }
        
        # Calculate next run time
        $nextRunTime = $startTime.AddHours($IntervalHours)
        Write-RunnerLog "Next execution scheduled for: $($nextRunTime.ToString('yyyy-MM-dd HH:mm:ss'))"
        
        # Sleep until next execution (check every minute for cancellation)
        while ((Get-Date) -lt $nextRunTime -and -not $cancelPressed) {
            Start-Sleep -Seconds 60
        }
        
        if ($cancelPressed) {
            Write-RunnerLog "Cancellation requested, exiting..." "WARN"
            break
        }
        
    } while (-not $cancelPressed)
}
catch {
    Write-RunnerLog "Critical error in continuous runner: $($_.Exception.Message)" "ERROR"
}
finally {
    Write-RunnerLog "=== CONTINUOUS RUNNER STOPPED ==="
}

<#
.SYNOPSIS
Runs the GroupQuery script continuously at specified intervals.

.DESCRIPTION
This script provides a continuous execution wrapper for the GroupQuery script.
It runs the script at regular intervals and includes logging and error handling.

.PARAMETER GroupQueryScriptPath
Full path to the GroupQuery.ps1 script (required)

.PARAMETER IntervalHours
Hours between executions (default: 1)

.PARAMETER LogPath
Path for the runner log file (optional, defaults to ContinuousRunnerLogs folder)

.PARAMETER ScriptParameters
Hashtable of parameters to pass to the GroupQuery script

.PARAMETER RunOnce
Run the script once and exit (useful for testing)

.EXAMPLE
# Basic usage - run every hour
.\ContinuousRunner.ps1 -GroupQueryScriptPath "C:\Scripts\GroupQuery.ps1"

.EXAMPLE
# Run every 2 hours with interactive auth
$params = @{
    UseInteractiveAuth = $true
    RunCleanupAfter = $false
}
.\ContinuousRunner.ps1 -GroupQueryScriptPath "C:\Scripts\GroupQuery.ps1" -IntervalHours 2 -ScriptParameters $params

.EXAMPLE
# Test run (execute once)
.\ContinuousRunner.ps1 -GroupQueryScriptPath "C:\Scripts\GroupQuery.ps1" -RunOnce

.NOTES
This script runs indefinitely until manually stopped (Ctrl+C) or the PowerShell session ends.
For production use, consider using Windows Task Scheduler instead.
#>
