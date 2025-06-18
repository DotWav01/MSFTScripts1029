# OneDrive Cleanup Script - Delete MP4 Transcript Files
# Reads OneDrive URLs from CSV and deletes .mp4 files containing "transcript" from recycle bins

param(
    [Parameter(Mandatory=$false)]
    [string]$TenantId,
    
    [Parameter(Mandatory=$false)]
    [string]$ClientId,
    
    [Parameter(Mandatory=$false)]
    [string]$Thumbprint,
    
    [Parameter(Mandatory=$false)]
    [string]$CsvFilePath,
    
    [Parameter(Mandatory=$false)]
    [string]$SharePointAdminUrl,
    
    [Parameter(Mandatory=$false)]
    [switch]$WhatIf,
    
    [Parameter(Mandatory=$false)]
    [string]$LogFilePath,
    
    [Parameter(Mandatory=$false)]
    [switch]$TestMode,
    
    [Parameter(Mandatory=$false)]
    [string]$TestOneDriveUrl,
    
    [Parameter(Mandatory=$false)]
    [string]$TestUserName,
    
    [Parameter(Mandatory=$false)]
    [switch]$UseInteractiveAuth,
    
    [Parameter(Mandatory=$false)]
    [string]$Username,
    
    [Parameter(Mandatory=$false)]
    [System.Security.SecureString]$Password,
    
    [Parameter(Mandatory=$false)]
    [string]$InteractiveClientId
)

#region Configuration - Update these values for your environment
$Config = @{
    TenantId            = ""              # Your Azure AD Tenant ID
    ClientId            = ""              # Your App Registration Client ID  
    Thumbprint          = "" # Certificate thumbprint for cert auth
    SharePointAdminUrl  = "" # Your SharePoint Admin Center URL
    InteractiveClientId = ""              # Client ID for interactive auth (can be same as ClientId)
}
#endregion Configuration

# Apply configuration defaults and simple defaults if parameters not provided
if ([string]::IsNullOrEmpty($TenantId)) { $TenantId = $Config.TenantId }
if ([string]::IsNullOrEmpty($ClientId)) { $ClientId = $Config.ClientId }
if ([string]::IsNullOrEmpty($Thumbprint)) { $Thumbprint = $Config.Thumbprint }
if ([string]::IsNullOrEmpty($SharePointAdminUrl)) { $SharePointAdminUrl = $Config.SharePointAdminUrl }
if ([string]::IsNullOrEmpty($InteractiveClientId)) { $InteractiveClientId = $Config.InteractiveClientId }
if ([string]::IsNullOrEmpty($CsvFilePath)) { $CsvFilePath = "OneDriveURLs.csv" }
if ([string]::IsNullOrEmpty($LogFilePath)) { $LogFilePath = "OneDriveCleanup_$(Get-Date -Format 'yyyyMMdd_HHmmss').log" }
if ([string]::IsNullOrEmpty($TestUserName)) { $TestUserName = "TestUser" }

# Function to write log entries
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Write-Host $logEntry
    Add-Content -Path $LogFilePath -Value $logEntry
}

# Function to validate CSV structure
function Test-CsvStructure {
    param([string]$CsvPath)
    
    if (!(Test-Path $CsvPath)) {
        throw "CSV file not found: $CsvPath"
    }
    
    $csvData = Import-Csv $CsvPath
    $requiredColumns = @("OneDriveURL")
    
    foreach ($column in $requiredColumns) {
        if ($column -notin $csvData[0].PSObject.Properties.Name) {
            throw "CSV missing required column: $column. Required columns: $($requiredColumns -join ', ')"
        }
    }
    
    return $csvData
}

# Function to connect using certificate authentication
function Connect-WithCertificate {
    param(
        [string]$Url,
        [string]$TenantId,
        [string]$ClientId,
        [string]$Thumbprint
    )
    
    Write-Log "Connecting with Certificate Authentication to: $Url"
    Connect-PnPOnline -Url $Url -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $TenantId -ErrorAction Stop
}

# Function to connect using interactive authentication
function Connect-WithInteractiveAuth {
    param(
        [string]$Url,
        [string]$Username = $null,
        [System.Security.SecureString]$Password = $null,
        [string]$ClientId = $null
    )
    
    if ($Username -and $Password) {
        Write-Log "Connecting with Username/Password Authentication to: $Url"
        $credential = New-Object System.Management.Automation.PSCredential($Username, $Password)
        Connect-PnPOnline -Url $Url -Credentials $credential -ErrorAction Stop
    } else {
        if ([string]::IsNullOrEmpty($ClientId)) {
            throw "InteractiveClientId is required for interactive browser authentication (since September 2024 PnP PowerShell security update)"
        }
        Write-Log "Connecting with Interactive Browser Authentication to: $Url"
        Write-Log "Using Client ID: $ClientId"
        Write-Log "You will be prompted to sign in through your browser..." "WARN"
        Connect-PnPOnline -Url $Url -ClientId $ClientId -Interactive -ErrorAction Stop
    }
}

# Function to validate authentication parameters
function Test-AuthenticationParameters {
    if ($UseInteractiveAuth) {
        Write-Log "Using Interactive Authentication mode"
        if ($Username -and !$Password) {
            throw "Password is required when Username is specified for interactive authentication"
        }
        if ($Password -and !$Username) {
            throw "Username is required when Password is specified for interactive authentication"
        }
        if (!$Username -and !$Password -and [string]::IsNullOrEmpty($InteractiveClientId)) {
            throw "InteractiveClientId is required for browser-based interactive authentication (PnP PowerShell security requirement since September 2024)"
        }
    } else {
        Write-Log "Using Certificate Authentication mode"
        # Validate required certificate parameters
        if ([string]::IsNullOrEmpty($TenantId)) {
            throw "TenantId is required for certificate authentication"
        }
        if ([string]::IsNullOrEmpty($ClientId)) {
            throw "ClientId is required for certificate authentication"
        }
        if ([string]::IsNullOrEmpty($Thumbprint)) {
            throw "Thumbprint is required for certificate authentication"
        }
        
        # Validate certificate is available
        Write-Log "Validating certificate with thumbprint: $Thumbprint"
        $cert = Get-ChildItem -Path "Cert:\CurrentUser\My\$Thumbprint" -ErrorAction SilentlyContinue
        if (-not $cert) {
            $cert = Get-ChildItem -Path "Cert:\LocalMachine\My\$Thumbprint" -ErrorAction SilentlyContinue
        }
        
        if (-not $cert) {
            throw "Certificate with thumbprint '$Thumbprint' not found in CurrentUser\My or LocalMachine\My certificate stores"
        }
        
        Write-Log "Certificate found: $($cert.Subject)"
    }
}

# Function to connect to SharePoint/OneDrive site
function Connect-ToSite {
    param([string]$Url)
    
    if ($UseInteractiveAuth) {
        Connect-WithInteractiveAuth -Url $Url -Username $Username -Password $Password -ClientId $InteractiveClientId
    } else {
        Connect-WithCertificate -Url $Url -TenantId $TenantId -ClientId $ClientId -Thumbprint $Thumbprint
    }
}

# Main script execution
try {
    Write-Log "=== ONEDRIVE CLEANUP SCRIPT STARTED ==="
    
    # Check if required modules are installed
    Write-Log "Checking required PowerShell modules..."
    $requiredModules = @("PnP.PowerShell")
    
    foreach ($module in $requiredModules) {
        if (!(Get-Module -ListAvailable -Name $module)) {
            Write-Log "Installing required module: $module" "WARN"
            Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser
        }
    }
    
    # Import required modules
    Write-Log "Importing PowerShell modules..."
    Import-Module PnP.PowerShell -Force
    
    # Determine operation mode
    if ($TestMode) {
        Write-Log "=== RUNNING IN TEST MODE ===" "WARN"
        
        if ([string]::IsNullOrEmpty($TestOneDriveUrl)) {
            $TestOneDriveUrl = Read-Host "Enter the OneDrive URL to test (e.g., https://contoso-my.sharepoint.com/personal/user_contoso_com)"
        }
        
        if ([string]::IsNullOrEmpty($TestOneDriveUrl) -or $TestOneDriveUrl -notlike "*-my.sharepoint.com/personal/*") {
            throw "Invalid OneDrive URL provided. Must be in format: https://tenant-my.sharepoint.com/personal/user_domain_com"
        }
        
        Write-Log "Test OneDrive URL: $TestOneDriveUrl"
        
        # Create a single test entry
        $oneDriveEntries = @([PSCustomObject]@{
            OneDriveURL = $TestOneDriveUrl
        })
        
        Write-Log "Test mode initialized with 1 OneDrive URL"
        
    } else {
        # CSV file mode
        Write-Log "CSV file mode enabled"
        Write-Log "Validating CSV file structure..."
        $csvData = Test-CsvStructure -CsvPath $CsvFilePath
        
        # Filter to only include entries with successful OneDrive URLs
        $oneDriveEntries = $csvData | Where-Object { 
            $_.OneDriveURL -and 
            $_.OneDriveURL -ne "" -and 
            (!$_.Status -or $_.Status -eq "Success") 
        }
        
        Write-Log "Found $($csvData.Count) total entries in CSV file"
        Write-Log "Found $($oneDriveEntries.Count) entries with valid OneDrive URLs for processing"
        
        if ($oneDriveEntries.Count -eq 0) {
            Write-Log "No valid OneDrive URLs found in CSV file. Exiting." "WARN"
            exit 0
        }
    }
    
    # Validate authentication parameters
    Test-AuthenticationParameters
    
    # Connect to SharePoint Online Admin
    Write-Log "Connecting to SharePoint Online Admin..."
    Connect-ToSite -Url $SharePointAdminUrl
    Write-Log "Successfully connected to SharePoint Online Admin"
    
    # Initialize counters
    $totalEntries = $oneDriveEntries.Count
    $processedEntries = 0
    $totalFilesFound = 0
    $totalFilesDeleted = 0
    $failedEntries = @()
    
    # Process each OneDrive URL
    foreach ($entry in $oneDriveEntries) {
        $processedEntries++
        $userOneDriveUrl = $entry.OneDriveURL.Trim()
        
        # Extract a display name from the URL for logging
        $displayName = "Unknown"
        if ($userOneDriveUrl -match "/personal/([^/]+)") {
            $displayName = $matches[1].Replace("_", "@").Replace("_", ".")
        }
        
        # Use UserPrincipalName if available in CSV
        if ($entry.UserPrincipalName) {
            $displayName = $entry.UserPrincipalName
        }
        
        Write-Log "Processing entry $processedEntries of $totalEntries`: $displayName"
        Write-Log "  OneDrive URL: $userOneDriveUrl"
        
        try {
            # Validate OneDrive URL format
            if ([string]::IsNullOrEmpty($userOneDriveUrl) -or $userOneDriveUrl -notlike "*-my.sharepoint.com/personal/*") {
                Write-Log "Invalid OneDrive URL format: '$userOneDriveUrl'" "WARN"
                $failedEntries += [PSCustomObject]@{
                    OneDriveURL = $userOneDriveUrl
                    Error = "Invalid OneDrive URL format"
                }
                continue
            }
            
            # Connect to the specific OneDrive site
            Write-Log "Connecting to OneDrive site..."
            Connect-ToSite -Url $userOneDriveUrl
            
            # Get recycle bin items
            Write-Log "Retrieving recycle bin items..."
            $recycleBinItems = Get-PnPRecycleBinItem -ErrorAction Stop
            
            # Filter for .mp4 files that contain "transcript" in the name
            $mp4Files = $recycleBinItems | Where-Object { 
                $_.LeafName -like "*.mp4" -and $_.LeafName -like "*transcript*"
            }
            $entryMp4Count = $mp4Files.Count
            $totalFilesFound += $entryMp4Count
            
            if ($entryMp4Count -eq 0) {
                Write-Log "No .mp4 transcript files found in recycle bin"
                continue
            }
            
            Write-Log "Found $entryMp4Count .mp4 transcript file(s) in recycle bin"
            
            # Log the files found
            foreach ($file in $mp4Files) {
                Write-Log "  - $($file.LeafName) (Deleted: $($file.DeletedDate), Size: $([math]::Round($file.Size/1MB, 2)) MB)"
            }
            
            if ($WhatIf) {
                Write-Log "[WHATIF] Would delete $entryMp4Count .mp4 transcript file(s)" "WARN"
                continue
            }
            
            # Delete the .mp4 transcript files
            Write-Log "Deleting .mp4 transcript files from recycle bin..."
            $entryDeletedCount = 0
            
            foreach ($file in $mp4Files) {
                try {
                    Clear-PnPRecycleBinItem -Identity $file.Id -Force -ErrorAction Stop
                    Write-Log "  ✓ Deleted: $($file.LeafName)"
                    $entryDeletedCount++
                    $totalFilesDeleted++
                }
                catch {
                    Write-Log "  ✗ Failed to delete: $($file.LeafName) - $($_.Exception.Message)" "ERROR"
                }
            }
            
            Write-Log "Completed processing: $entryDeletedCount of $entryMp4Count files deleted"
            
        }
        catch {
            $errorMessage = $_.Exception.Message
            Write-Log "Error processing OneDrive $userOneDriveUrl`: $errorMessage" "ERROR"
            $failedEntries += [PSCustomObject]@{
                OneDriveURL = $userOneDriveUrl
                Error = $errorMessage
            }
        }
    }
    
    # Summary report
    Write-Log "=== CLEANUP SUMMARY REPORT ===" "INFO"
    Write-Log "Authentication method: $(if ($UseInteractiveAuth) { 'Interactive' } else { 'Certificate' })"
    if ($TestMode) {
        Write-Log "Source: Test mode"
    } else {
        Write-Log "Source: CSV file '$CsvFilePath'"
    }
    Write-Log "Total OneDrive URLs processed: $processedEntries"
    Write-Log "Total .mp4 transcript files found: $totalFilesFound"
    
    if ($WhatIf) {
        Write-Log "Total .mp4 transcript files that would be deleted: $totalFilesFound (WHATIF MODE)" "WARN"
    } else {
        Write-Log "Total .mp4 transcript files deleted: $totalFilesDeleted"
    }
    
    if ($failedEntries.Count -gt 0) {
        Write-Log "Failed entries: $($failedEntries.Count)" "WARN"
        foreach ($failed in $failedEntries) {
            Write-Log "  - $($failed.OneDriveURL): $($failed.Error)" "WARN"
        }
    }
    
    Write-Log "Log file saved to: $LogFilePath"
    
}
catch {
    Write-Log "Critical error: $($_.Exception.Message)" "ERROR"
    throw
}
finally {
    # Disconnect from SharePoint Online
    try {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        Write-Log "Disconnected from SharePoint Online"
    }
    catch {
        # Ignore disconnect errors
    }
}

<#
.SYNOPSIS
Deletes .mp4 transcript files from OneDrive recycle bins using URLs from CSV file.

.DESCRIPTION
This script reads OneDrive URLs from a CSV file and deletes .mp4 files that contain "transcript" 
in the filename from their recycle bins. It supports both Azure App Registration certificate 
authentication and interactive user authentication.

.PARAMETER TenantId
Your Azure AD Tenant ID (defaults to configured value, required for certificate authentication)

.PARAMETER ClientId
The Client ID of your Azure App Registration (defaults to configured value, required for certificate authentication)

.PARAMETER Thumbprint
The certificate thumbprint for your Azure App Registration (defaults to configured value, required for certificate authentication)

.PARAMETER CsvFilePath
Path to CSV file containing OneDrive URLs (defaults to OneDriveURLs.csv)

.PARAMETER SharePointAdminUrl
Your SharePoint Admin Center URL (defaults to configured value)

.PARAMETER WhatIf
Preview mode - shows what would be deleted without actually deleting

.PARAMETER LogFilePath
Path for log file (optional, defaults to timestamped file in current directory)

.PARAMETER TestMode
Enables test mode where you can manually specify a OneDrive URL

.PARAMETER TestOneDriveUrl
The OneDrive URL to test (only used in test mode)

.PARAMETER UseInteractiveAuth
Use interactive authentication instead of certificate authentication

.PARAMETER Username
Username for interactive authentication (optional)

.PARAMETER Password
SecureString password for interactive authentication (optional)

.PARAMETER InteractiveClientId
The Client ID for interactive browser authentication (defaults to configured value, required for browser sign-in)

.EXAMPLE
# Certificate authentication (using configured values)
.\OneDrive-Cleanup.ps1 -WhatIf

.EXAMPLE
# Interactive authentication with browser sign-in (using configured values)
.\OneDrive-Cleanup.ps1 -UseInteractiveAuth -WhatIf

.EXAMPLE
# Interactive authentication with username/password
$securePassword = ConvertTo-SecureString "YourPassword" -AsPlainText -Force
.\OneDrive-Cleanup.ps1 -UseInteractiveAuth -Username "user@contoso.com" -Password $securePassword

.EXAMPLE
# Test mode with single OneDrive
.\OneDrive-Cleanup.ps1 -TestMode -UseInteractiveAuth -WhatIf

.EXAMPLE
# Use custom CSV file and override SharePoint URL
.\OneDrive-Cleanup.ps1 -CsvFilePath "C:\Scripts\MyOneDriveList.csv" -SharePointAdminUrl "https://different-admin.sharepoint.com"

.NOTES
CONFIGURATION: Update the configuration section at the top of this script with your tenant values:
- TenantId: Your Azure AD Tenant ID
- ClientId: Your App Registration Client ID  
- Thumbprint: Your certificate thumbprint
- SharePointAdminUrl: Your SharePoint Admin Center URL
- InteractiveClientId: Client ID for interactive auth (can be same as ClientId)

CSV File Format:
The script expects a CSV with at least an 'OneDriveURL' column. 
Optional columns: UserPrincipalName, DisplayName, Status

Authentication Options:
1. Certificate Authentication (for unattended execution):
   - Requires TenantId, ClientId, and Thumbprint parameters
   - App Registration Requirements: Sites.FullControl.All

2. Interactive Authentication (for testing/manual execution):
   - Use -UseInteractiveAuth switch
   - Optional Username/Password parameters (if not provided, browser sign-in will be used)
   - User must have appropriate permissions to access OneDrive sites

Module Requirements:
- PnP.PowerShell

Permissions Required:
- SharePoint Administrator role or equivalent permissions to access OneDrive sites
- For certificate auth: App registration with Sites.FullControl.All permissions
#>
