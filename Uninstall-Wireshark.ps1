# Uninstall-Applications.ps1
# PowerShell script to uninstall Wireshark and Npcap applications
# Designed for Intune Win32 App deployment

param(
    [string]$LogPath = "$env:ProgramData\IntuneApps\UninstallLogs"
)

# Create log directory if it doesn't exist
if (!(Test-Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
}

$LogFile = Join-Path $LogPath "AppUninstall_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Function to write to log file and console
function Write-LogMessage {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Write-Output $logEntry
    Add-Content -Path $LogFile -Value $logEntry
}

# Function to get installed application from registry
function Get-InstalledApplication {
    param([string]$DisplayName)
    
    $UninstallKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    foreach ($Key in $UninstallKeys) {
        try {
            $Apps = Get-ItemProperty $Key -ErrorAction SilentlyContinue | Where-Object {
                $_.DisplayName -like "*$DisplayName*" -and $_.UninstallString
            }
            if ($Apps) {
                return $Apps | Select-Object -First 1
            }
        }
        catch {
            # Continue to next registry path
        }
    }
    return $null
}

# Function to uninstall using Windows Installer (MSI)
function Uninstall-MSIApplication {
    param(
        [string]$ProductCode,
        [string]$AppName
    )
    
    try {
        Write-LogMessage "Uninstalling $AppName using MSI product code: $ProductCode" "INFO"
        
        $Arguments = "/x `"$ProductCode`" /quiet /norestart"
        $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $Arguments -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 1605) {
            Write-LogMessage "$AppName uninstalled successfully (Exit Code: $($process.ExitCode))" "SUCCESS"
            return $true
        } else {
            Write-LogMessage "$AppName uninstall failed (Exit Code: $($process.ExitCode))" "ERROR"
            return $false
        }
    }
    catch {
        Write-LogMessage "Exception during $AppName MSI uninstall: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Function to uninstall using EXE uninstaller
function Uninstall-EXEApplication {
    param(
        [string]$UninstallString,
        [string]$AppName,
        [string]$SilentArgs = ""
    )
    
    try {
        Write-LogMessage "Uninstalling $AppName using EXE uninstaller" "INFO"
        Write-LogMessage "Uninstall string: $UninstallString" "INFO"
        
        # Parse the uninstall string to separate executable and arguments
        if ($UninstallString -match '^"([^"]+)"(.*)$') {
            $ExePath = $matches[1]
            $ExistingArgs = $matches[2].Trim()
        } else {
            $Parts = $UninstallString -split ' ', 2
            $ExePath = $Parts[0].Trim('"')
            $ExistingArgs = if ($Parts.Count -gt 1) { $Parts[1] } else { "" }
        }
        
        # Combine existing arguments with silent arguments
        $Arguments = "$ExistingArgs $SilentArgs".Trim()
        
        Write-LogMessage "Executable: $ExePath" "INFO"
        Write-LogMessage "Arguments: $Arguments" "INFO"
        
        if (!(Test-Path $ExePath)) {
            Write-LogMessage "Uninstaller not found: $ExePath" "ERROR"
            return $false
        }
        
        $process = Start-Process -FilePath $ExePath -ArgumentList $Arguments -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -eq 0) {
            Write-LogMessage "$AppName uninstalled successfully (Exit Code: $($process.ExitCode))" "SUCCESS"
            return $true
        } else {
            Write-LogMessage "$AppName uninstall failed (Exit Code: $($process.ExitCode))" "ERROR"
            return $false
        }
    }
    catch {
        Write-LogMessage "Exception during $AppName EXE uninstall: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Function to uninstall application
function Uninstall-Application {
    param(
        [string]$DisplayName,
        [string]$SilentArgs = ""
    )
    
    Write-LogMessage "Looking for installed application: $DisplayName" "INFO"
    
    $App = Get-InstalledApplication -DisplayName $DisplayName
    
    if (-not $App) {
        Write-LogMessage "$DisplayName not found in installed programs" "WARNING"
        return $true  # Consider as success if app is not installed
    }
    
    Write-LogMessage "Found $DisplayName - Version: $($App.DisplayVersion)" "INFO"
    
    # Check if it's an MSI installation (has ProductCode)
    if ($App.PSObject.Properties.Name -contains "PSChildName" -and $App.PSChildName -match "^{[0-9A-F-]{36}}$") {
        return Uninstall-MSIApplication -ProductCode $App.PSChildName -AppName $DisplayName
    }
    # Check for MSI in UninstallString
    elseif ($App.UninstallString -match "msiexec") {
        # Extract product code from MSI uninstall string
        if ($App.UninstallString -match "{[0-9A-F-]{36}}") {
            $ProductCode = $matches[0]
            return Uninstall-MSIApplication -ProductCode $ProductCode -AppName $DisplayName
        }
    }
    
    # Use EXE uninstaller
    return Uninstall-EXEApplication -UninstallString $App.UninstallString -AppName $DisplayName -SilentArgs $SilentArgs
}

# Main uninstall logic
try {
    Write-LogMessage "=== Starting Application Uninstall Process ===" "INFO"
    
    # Define applications to uninstall in reverse order (Wireshark first, then Npcap)
    $Applications = @(
        @{
            Name = "Wireshark"
            SilentArgs = "/S"
            Priority = 1
        },
        @{
            Name = "Npcap"
            SilentArgs = "/S"
            Priority = 2
        }
    )
    
    $UninstallResults = @()
    
    # Sort applications by priority
    $Applications = $Applications | Sort-Object Priority
    
    foreach ($App in $Applications) {
        Write-LogMessage "Starting uninstall of $($App.Name)..." "INFO"
        
        $UninstallSuccess = Uninstall-Application -DisplayName $App.Name -SilentArgs $App.SilentArgs
        
        $UninstallResults += @{
            App = $App.Name
            Success = $UninstallSuccess
        }
        
        if ($UninstallSuccess) {
            Write-LogMessage "$($App.Name) uninstall completed successfully" "SUCCESS"
            # Add delay between uninstalls to ensure clean removal
            Start-Sleep -Seconds 10
        } else {
            Write-LogMessage "$($App.Name) uninstall encountered issues" "WARNING"
            # Continue with next application even if current one fails
            Start-Sleep -Seconds 5
        }
    }
    
    # Summary
    Write-LogMessage "=== Uninstall Summary ===" "INFO"
    $SuccessCount = 0
    foreach ($Result in $UninstallResults) {
        $Status = if ($Result.Success) { "SUCCESS" } else { "FAILED" }
        Write-LogMessage "$($Result.App): $Status" "INFO"
        if ($Result.Success) { $SuccessCount++ }
    }
    
    Write-LogMessage "Successfully uninstalled $SuccessCount of $($UninstallResults.Count) applications" "INFO"
    
    # Clean up application data folders (optional)
    Write-LogMessage "Cleaning up remaining application data..." "INFO"
    
    $CleanupPaths = @(
        "$env:ProgramData\Wireshark",
        "$env:APPDATA\Wireshark",
        "$env:LOCALAPPDATA\Wireshark"
    )
    
    foreach ($Path in $CleanupPaths) {
        if (Test-Path $Path) {
            try {
                Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
                Write-LogMessage "Removed folder: $Path" "INFO"
            }
            catch {
                Write-LogMessage "Could not remove folder $Path : $($_.Exception.Message)" "WARNING"
            }
        }
    }
    
    # Set exit code (always return 0 for uninstall unless critical error)
    Write-LogMessage "Uninstall process completed" "SUCCESS"
    exit 0
}
catch {
    Write-LogMessage "Critical error in uninstall script: $($_.Exception.Message)" "ERROR"
    Write-LogMessage "Stack trace: $($_.ScriptStackTrace)" "ERROR"
    exit 1
}
finally {
    Write-LogMessage "=== Uninstall Process Completed ===" "INFO"
}
