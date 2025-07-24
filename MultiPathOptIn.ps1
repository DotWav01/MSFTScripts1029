# Define the registry path and property details
$registryPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\RdpCloudStackSettings"
$propertyName = "SmilesV3ActivationThreshold"
$propertyValue = 100

# Define log file path
$logPath = "C:\temp"
$logFile = Join-Path $logPath "RdpCloudStackSettings_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Function to write to both console and log file
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] $Message"
    Write-Host $logEntry
    Add-Content -Path $logFile -Value $logEntry
}

# Create log directory if it doesn't exist
if (!(Test-Path $logPath)) {
    New-Item -Path $logPath -ItemType Directory -Force | Out-Null
}

# Initialize log file
Write-Log "Starting RdpCloudStackSettings registry configuration script"
Write-Log "Target registry path: $registryPath"
Write-Log "Property name: $propertyName"
Write-Log "Property value: $propertyValue"

# Check if the registry key exists, create it if it doesn't
if (!(Test-Path $registryPath)) {
    Write-Log "Registry key does not exist. Creating: $registryPath"
    try {
        New-Item -Path $registryPath -Force | Out-Null
        Write-Log "Registry key created successfully"
    }
    catch {
        Write-Log "ERROR: Failed to create registry key - $($_.Exception.Message)"
        exit 1
    }
} else {
    Write-Log "Registry key already exists: $registryPath"
}

# Check if the registry property already exists
try {
    $existingValue = Get-ItemProperty -Path $registryPath -Name $propertyName -ErrorAction Stop
    Write-Log "Registry setting '$propertyName' already exists with value: $($existingValue.$propertyName)"
    Write-Log "Exiting without making changes"
    Write-Log "Script completed - no changes made"
    exit 0
}
catch {
    # Property doesn't exist, so create it
    Write-Log "Registry setting '$propertyName' does not exist. Creating it with value: $propertyValue"
    try {
        New-ItemProperty -Path $registryPath -Name $propertyName -Value $propertyValue -PropertyType DWord -Force | Out-Null
        Write-Log "Registry setting created successfully"
        Write-Log "Script completed - registry setting created"
    }
    catch {
        Write-Log "ERROR: Failed to create registry property - $($_.Exception.Message)"
        exit 1
    }
}

Write-Log "Log file saved to: $logFile"
