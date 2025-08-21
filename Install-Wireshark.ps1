# Install-Applications.ps1
# PowerShell script to install EXE applications from the same directory
# Designed for Intune Win32 App deployment

param(
    [string]$LogPath = "$env:ProgramData\IntuneApps\InstallLogs"
)

# Create log directory if it doesn't exist
if (!(Test-Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
}

$LogFile = Join-Path $LogPath "AppInstall_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Function to write to log file and console
function Write-LogMessage {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Write-Output $logEntry
    Add-Content -Path $LogFile -Value $logEntry
}

# Function to install an EXE file
function Install-Application {
    param(
        [string]$ExePath,
        [string]$Arguments = "",
        [string]$AppName
    )
    
    try {
        Write-LogMessage "Starting installation of $AppName" "INFO"
        Write-LogMessage "Executable path: $ExePath" "INFO"
        Write-LogMessage "Arguments: $Arguments" "INFO"
        
        # Check if file exists
        if (!(Test-Path $ExePath)) {
            Write-LogMessage "File not found: $ExePath" "ERROR"
            return $false
        }
        
        # Start the installation process
        $process = Start-Process -FilePath $ExePath -ArgumentList $Arguments -Wait -PassThru -NoNewWindow
        
        # Check exit code
        if ($process.ExitCode -eq 0) {
            Write-LogMessage "$AppName installed successfully (Exit Code: $($process.ExitCode))" "SUCCESS"
            return $true
        } else {
            Write-LogMessage "$AppName installation failed (Exit Code: $($process.ExitCode))" "ERROR"
            return $false
        }
    }
    catch {
        Write-LogMessage "Exception occurred during $AppName installation: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Main installation logic
try {
    Write-LogMessage "=== Starting Application Installation Process ===" "INFO"
    
    # Get the directory where this script is located
    $ScriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
    Write-LogMessage "Script directory: $ScriptDirectory" "INFO"
    
    # Define applications to install in specific order (Npcap must be installed before Wireshark)
    $Applications = @(
        @{
            Name = "Npcap"
            Pattern = "*npcap*.exe"
            Arguments = "/S"  # Silent install for Npcap
            Priority = 1
        },
        @{
            Name = "Wireshark" 
            Pattern = "*wireshark*.exe"
            Arguments = "/S"  # Silent install for Wireshark
            Priority = 2
            DependsOn = "Npcap"
        }
    )
    
    $InstallResults = @()
    
    # Sort applications by priority to ensure correct installation order
    $Applications = $Applications | Sort-Object Priority
    
    foreach ($App in $Applications) {
        Write-LogMessage "Looking for $($App.Name) installer..." "INFO"
        
        # Check if dependency was successfully installed (if applicable)
        if ($App.DependsOn) {
            $DependencyResult = $InstallResults | Where-Object { $_.App -eq $App.DependsOn }
            if ($DependencyResult -and -not $DependencyResult.Success) {
                Write-LogMessage "Skipping $($App.Name) installation because dependency $($App.DependsOn) failed to install" "ERROR"
                $InstallResults += @{
                    App = $App.Name
                    Success = $false
                    Reason = "Dependency $($App.DependsOn) installation failed"
                }
                continue
            }
            elseif (-not $DependencyResult) {
                Write-LogMessage "Skipping $($App.Name) installation because dependency $($App.DependsOn) was not found/attempted" "ERROR"
                $InstallResults += @{
                    App = $App.Name
                    Success = $false
                    Reason = "Dependency $($App.DependsOn) not found"
                }
                continue
            }
        }
        
        # Find the EXE file matching the pattern
        $ExeFiles = Get-ChildItem -Path $ScriptDirectory -Filter $App.Pattern -File
        
        if ($ExeFiles.Count -eq 0) {
            Write-LogMessage "No $($App.Name) installer found matching pattern: $($App.Pattern)" "WARNING"
            $InstallResults += @{
                App = $App.Name
                Success = $false
                Reason = "Installer not found"
            }
            continue
        }
        
        if ($ExeFiles.Count -gt 1) {
            Write-LogMessage "Multiple $($App.Name) installers found. Using the first one: $($ExeFiles[0].Name)" "WARNING"
        }
        
        $ExeFile = $ExeFiles[0]
        $ExePath = $ExeFile.FullName
        
        # Install the application and wait for it to complete
        Write-LogMessage "Installing $($App.Name) - waiting for completion..." "INFO"
        $InstallSuccess = Install-Application -ExePath $ExePath -Arguments $App.Arguments -AppName $App.Name
        
        $InstallResults += @{
            App = $App.Name
            Success = $InstallSuccess
            Path = $ExePath
        }
        
        if ($InstallSuccess) {
            Write-LogMessage "$($App.Name) installation completed successfully. Ready for next application." "SUCCESS"
            # Add a longer delay after successful installation to ensure system is ready
            Start-Sleep -Seconds 10
        } else {
            Write-LogMessage "$($App.Name) installation failed. Stopping installation process." "ERROR"
            break  # Stop installing subsequent applications if current one fails
        }
    }
    
    # Summary
    Write-LogMessage "=== Installation Summary ===" "INFO"
    $SuccessCount = 0
    foreach ($Result in $InstallResults) {
        $Status = if ($Result.Success) { "SUCCESS" } else { "FAILED" }
        Write-LogMessage "$($Result.App): $Status" "INFO"
        if ($Result.Success) { $SuccessCount++ }
    }
    
    Write-LogMessage "Successfully installed $SuccessCount of $($InstallResults.Count) applications" "INFO"
    
    # Set exit code based on results
    if ($SuccessCount -eq $InstallResults.Count) {
        Write-LogMessage "All installations completed successfully" "SUCCESS"
        exit 0
    } else {
        Write-LogMessage "One or more installations failed" "ERROR"
        exit 1
    }
}
catch {
    Write-LogMessage "Critical error in main script: $($_.Exception.Message)" "ERROR"
    Write-LogMessage "Stack trace: $($_.ScriptStackTrace)" "ERROR"
    exit 1
}
finally {
    Write-LogMessage "=== Installation Process Completed ===" "INFO"
}
