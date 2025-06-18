# Group Query Script - Get OneDrive URLs from Entra ID Group
# Queries an Entra ID group and updates/creates CSV with member OneDrive URLs

param(
    [Parameter(Mandatory=$false)]
    [string[]]$GroupNames,
    
    [Parameter(Mandatory=$false)]
    [string]$GroupName,
    
    [Parameter(Mandatory=$false)]
    [string[]]$UserList,
    
    [Parameter(Mandatory=$false)]
    [string]$TenantId,
    
    [Parameter(Mandatory=$false)]
    [string]$ClientId,
    
    [Parameter(Mandatory=$false)]
    [string]$Thumbprint,
    
    [Parameter(Mandatory=$false)]
    [switch]$UseInteractiveAuth,
    
    [Parameter(Mandatory=$false)]
    [string]$Username,
    
    [Parameter(Mandatory=$false)]
    [System.Security.SecureString]$Password,
    
    [Parameter(Mandatory=$false)]
    [string]$CsvFilePath,
    
    [Parameter(Mandatory=$false)]
    [string]$LogFilePath,
    
    [Parameter(Mandatory=$false)]
    [string]$CleanupScriptPath,
    
    [Parameter(Mandatory=$false)]
    [string]$SharePointAdminUrl,
    
    [Parameter(Mandatory=$false)]
    [switch]$RunCleanupAfter,
    
    [Parameter(Mandatory=$false)]
    [switch]$WhatIfCleanup
)

#region Configuration - Update these values for your environment
$Config = @{
    TenantId            = ""              # Your Azure AD Tenant ID
    ClientId            = ""              # Your App Registration Client ID  
    Thumbprint          = "" # Certificate thumbprint for cert auth
    SharePointAdminUrl  = "" # Your SharePoint Admin Center URL
    
    # Optional: Hard-code the group name(s) to always query the same group(s)
    DefaultGroupNames   = @("","")       # Default group names if not provided as parameter
    
    # Optional: Hard-code specific users instead of querying groups
    # If this array has values, it will override group query
    # Format: User Principal Names (email addresses)
    HardCodedUsers = @(
        # "user1@yourdomain.com",
        # "user2@yourdomain.com",
        # "user3@yourdomain.com"
    )
}
#endregion Configuration

# Create GroupQueryLogs folder if it doesn't exist (next to the script file)
$ScriptDirectory = Split-Path $MyInvocation.MyCommand.Path
$LogsFolder = Join-Path $ScriptDirectory "GroupQueryLogs"
if (!(Test-Path $LogsFolder)) {
    New-Item -ItemType Directory -Path $LogsFolder -Force | Out-Null
}

# Apply configuration defaults and simple defaults if parameters not provided
if ([string]::IsNullOrEmpty($TenantId)) { $TenantId = $Config.TenantId }
if ([string]::IsNullOrEmpty($ClientId)) { $ClientId = $Config.ClientId }
if ([string]::IsNullOrEmpty($Thumbprint)) { $Thumbprint = $Config.Thumbprint }
if ([string]::IsNullOrEmpty($SharePointAdminUrl)) { $SharePointAdminUrl = $Config.SharePointAdminUrl }
if (!$UserList -or $UserList.Count -eq 0) { $UserList = $Config.HardCodedUsers }
if ([string]::IsNullOrEmpty($CsvFilePath)) { $CsvFilePath = "OneDriveURLs.csv" }
if ([string]::IsNullOrEmpty($LogFilePath)) { 
    $LogFilePath = Join-Path $LogsFolder "GroupQuery_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
}

# Handle group name(s) - support both single and multiple groups
if ($GroupName -and !$GroupNames) {
    # Convert single GroupName to GroupNames array for consistency
    $GroupNames = @($GroupName)
} elseif (!$GroupNames -or $GroupNames.Count -eq 0) {
    # Use configured default group names
    $GroupNames = $Config.DefaultGroupNames
}

# Set RunCleanupAfter to true by default (can be overridden with -RunCleanupAfter:$false)
if (!$PSBoundParameters.ContainsKey('RunCleanupAfter')) { $RunCleanupAfter = $true }

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

# Function to connect to Microsoft Graph
function Connect-ToMicrosoftGraph {
    Write-Log "Connecting to Microsoft Graph..."
    
    if ($UseInteractiveAuth) {
        if ($Username -and $Password) {
            Write-Log "Connecting with Username/Password Authentication"
            $credential = New-Object System.Management.Automation.PSCredential($Username, $Password)
            Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -Credential $credential -ErrorAction Stop
        } else {
            Write-Log "Connecting with Interactive Browser Authentication"
            Write-Log "Using Client ID: $ClientId"
            Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -ErrorAction Stop
        }
    } else {
        # Certificate authentication
        if ([string]::IsNullOrEmpty($Thumbprint)) {
            throw "Thumbprint is required for certificate authentication"
        }
        Write-Log "Connecting with Certificate Authentication"
        Write-Log "Using Client ID: $ClientId, Tenant ID: $TenantId"
        Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $Thumbprint -ErrorAction Stop
    }
    
    Write-Log "Successfully connected to Microsoft Graph"
}

# Function to validate certificate if using certificate auth
function Test-CertificateAuth {
    if (!$UseInteractiveAuth) {
        if ([string]::IsNullOrEmpty($Thumbprint)) {
            throw "Thumbprint is required for certificate authentication"
        }
        
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

# Function to load existing CSV data
function Get-ExistingCsvData {
    param([string]$CsvPath)
    
    if (Test-Path $CsvPath) {
        Write-Log "Loading existing CSV file: $CsvPath"
        $existingData = Import-Csv $CsvPath
        Write-Log "Found $($existingData.Count) existing entries in CSV"
        return $existingData
    } else {
        Write-Log "CSV file does not exist, will create new file: $CsvPath"
        return @()
    }
}

# Function to process multiple groups and get OneDrive URLs
function Get-OneDriveUrlsFromGroups {
    param(
        [string[]]$GroupNames,
        [string]$OutputPath
    )
    
    Write-Log "=== QUERYING MULTIPLE ENTRA ID GROUPS FOR ONEDRIVE URLS ==="
    Write-Log "Groups to process: $($GroupNames -join ', ')"
    
    # Initialize results array for all groups
    $allResults = @()
    $groupCounter = 0
    
    foreach ($groupName in $GroupNames) {
        $groupCounter++
        Write-Log "Processing group $groupCounter of $($GroupNames.Count): $groupName"
        
        try {
            # Find the group by name
            Write-Log "  Searching for group: $groupName"
            $group = Get-MgGroup -Filter "displayName eq '$groupName'" -ErrorAction Stop
            
            if (!$group) {
                Write-Log "  Group '$groupName' not found" "WARN"
                continue
            }
            
            if ($group.Count -gt 1) {
                Write-Log "  Multiple groups found with name '$groupName'. Using the first one." "WARN"
                $group = $group[0]
            }
            
            Write-Log "  Found group: $($group.DisplayName) (ID: $($group.Id))"
            
            # Get group members
            Write-Log "  Retrieving group members..."
            $groupMembers = Get-MgGroupMember -GroupId $group.Id -All
            
            # Filter to only get user members
            $users = $groupMembers | Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.user' }
            
            if ($users.Count -eq 0) {
                Write-Log "  No user members found in group '$groupName'" "WARN"
                continue
            }
            
            Write-Log "  Found $($users.Count) user members in the group"
            
            # Process users in this group
            $userCounter = 0
            foreach ($userMember in $users) {
                $userCounter++
                $userId = $userMember.Id
                
                Write-Log "  Processing user $userCounter of $($users.Count) from group '$groupName'..."
                
                try {
                    # Get full user information
                    $mgUser = Get-MgUser -UserId $userId -ErrorAction Stop
                    $userPrincipalName = $mgUser.UserPrincipalName
                    
                    Write-Log "    User: $userPrincipalName"
                    
                    # Check if user already processed (avoid duplicates across groups)
                    $existingUser = $allResults | Where-Object { $_.UserPrincipalName -eq $userPrincipalName }
                    if ($existingUser) {
                        Write-Log "    User already processed from another group, skipping duplicate"
                        continue
                    }
                    
                    # Get user's default OneDrive
                    $oneDrive = Get-MgUserDefaultDrive -UserId $mgUser.Id -ErrorAction Stop
                    
                    if ($oneDrive) {
                        # Remove /Documents from the end of the WebUrl to get the base OneDrive URL
                        $baseUrl = $oneDrive.WebUrl
                        if ($baseUrl.EndsWith("/Documents")) {
                            $oneDriveUrl = $baseUrl.Substring(0, $baseUrl.Length - 10)
                        } else {
                            $oneDriveUrl = $baseUrl
                        }
                        $status = "Success"
                        $errorMsg = ""
                        Write-Log "    OneDrive URL found: $oneDriveUrl"
                    } else {
                        $oneDriveUrl = ""
                        $status = "No OneDrive"
                        $errorMsg = "OneDrive not found for user"
                        Write-Log "    OneDrive not found" "WARN"
                    }
                }
                catch {
                    $userPrincipalName = if ($mgUser) { $mgUser.UserPrincipalName } else { "Unknown" }
                    $oneDriveUrl = ""
                    $status = "Error"
                    $errorMsg = $_.Exception.Message
                    Write-Log "    Error: $errorMsg" "ERROR"
                }
                
                # Add result to array with group information
                $allResults += [PSCustomObject]@{
                    UserPrincipalName = $userPrincipalName
                    DisplayName = if ($mgUser) { $mgUser.DisplayName } else { "N/A" }
                    OneDriveURL = $oneDriveUrl
                    Status = $status
                    ErrorMessage = $errorMsg
                    SourceGroup = $groupName
                }
                
                # Add small delay to avoid throttling
                Start-Sleep -Milliseconds 200
            }
            
            Write-Log "  Completed processing group '$groupName'"
        }
        catch {
            Write-Log "  Error processing group '$groupName': $($_.Exception.Message)" "ERROR"
        }
    }
    
    return $allResults
}
function Get-OneDriveUrlsFromUserList {
    param(
        [string[]]$UserList,
        [string]$OutputPath
    )
    
    Write-Log "=== PROCESSING HARD-CODED USER LIST ==="
    Write-Log "Found $($UserList.Count) users in hard-coded list"
    
    # Initialize results array
    $newResults = @()
    $counter = 0
    
    foreach ($userPrincipalName in $UserList) {
        $counter++
        
        Write-Log "Processing user $counter of $($UserList.Count): $userPrincipalName"
        
        try {
            # Get full user information
            $mgUser = Get-MgUser -UserId $userPrincipalName -ErrorAction Stop
            
            Write-Log "  User found: $($mgUser.DisplayName)"
            
            # Get user's default OneDrive
            $oneDrive = Get-MgUserDefaultDrive -UserId $mgUser.Id -ErrorAction Stop
            
            if ($oneDrive) {
                # Remove /Documents from the end of the WebUrl to get the base OneDrive URL
                $baseUrl = $oneDrive.WebUrl
                if ($baseUrl.EndsWith("/Documents")) {
                    $oneDriveUrl = $baseUrl.Substring(0, $baseUrl.Length - 10)
                } else {
                    $oneDriveUrl = $baseUrl
                }
                $status = "Success"
                $errorMsg = ""
                Write-Log "  OneDrive URL found: $oneDriveUrl"
            } else {
                $oneDriveUrl = ""
                $status = "No OneDrive"
                $errorMsg = "OneDrive not found for user"
                Write-Log "  OneDrive not found" "WARN"
            }
        }
        catch {
            $oneDriveUrl = ""
            $status = "Error"
            $errorMsg = $_.Exception.Message
            Write-Log "  Error: $errorMsg" "ERROR"
        }
        
        # Add result to array
        $newResults += [PSCustomObject]@{
            UserPrincipalName = $userPrincipalName
            DisplayName = if ($mgUser) { $mgUser.DisplayName } else { "N/A" }
            OneDriveURL = $oneDriveUrl
            Status = $status
            ErrorMessage = $errorMsg
        }
        
        # Add small delay to avoid throttling
        Start-Sleep -Milliseconds 200
    }
    
    return $newResults
}
function Merge-CsvData {
    param(
        [array]$ExistingData,
        [array]$NewResults
    )
    
    Write-Log "Merging new results with existing data..."
    
    # Create hashtable of existing data for quick lookup
    $existingHash = @{}
    foreach ($item in $ExistingData) {
        if ($item.UserPrincipalName) {
            $existingHash[$item.UserPrincipalName] = $item
        }
    }
    
    # Process new results
    $mergedData = @()
    $addedCount = 0
    $updatedCount = 0
    
    foreach ($newItem in $NewResults) {
        if ($existingHash.ContainsKey($newItem.UserPrincipalName)) {
            # Update existing entry
            $existingItem = $existingHash[$newItem.UserPrincipalName]
            $existingItem.DisplayName = $newItem.DisplayName
            $existingItem.OneDriveURL = $newItem.OneDriveURL
            $existingItem.Status = $newItem.Status
            $existingItem.ErrorMessage = $newItem.ErrorMessage
            $existingItem.LastUpdated = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $mergedData += $existingItem
            $updatedCount++
        } else {
            # Add new entry
            $newItem | Add-Member -NotePropertyName "LastUpdated" -NotePropertyValue (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            $mergedData += $newItem
            $addedCount++
        }
        # Remove from existing hash so we know what's left
        $existingHash.Remove($newItem.UserPrincipalName)
    }
    
    # Add remaining existing entries that weren't in the new results
    foreach ($remainingItem in $existingHash.Values) {
        $mergedData += $remainingItem
    }
    
    Write-Log "Merge complete - Added: $addedCount, Updated: $updatedCount, Total: $($mergedData.Count)"
    return $mergedData
}

# Main script execution
try {
    Write-Log "=== GROUP QUERY SCRIPT STARTED ==="
    Write-Log "Log file location: $LogFilePath"
    
    # Determine operation mode
    $useHardCodedUsers = $UserList -and $UserList.Count -gt 0
    $useGroupQuery = $GroupNames -and $GroupNames.Count -gt 0
    
    if ($useHardCodedUsers -and $useGroupQuery) {
        Write-Log "Both UserList and GroupNames provided. Using hard-coded UserList and ignoring GroupNames." "WARN"
        $useGroupQuery = $false
    }
    
    if (!$useHardCodedUsers -and !$useGroupQuery) {
        throw "Either GroupNames parameter or hard-coded UserList in configuration must be provided"
    }
    
    if ($useHardCodedUsers) {
        Write-Log "Hard-coded user list mode enabled"
        Write-Log "CSV Output: $CsvFilePath"
        Write-Log "Users to process: $($UserList -join ', ')"
    } else {
        Write-Log "Group query mode enabled"
        Write-Log "Groups: $($GroupNames -join ', ')"
        Write-Log "CSV Output: $CsvFilePath"
    }
    
    # Validate authentication parameters
    Test-CertificateAuth
    
    # Check if Microsoft Graph modules are installed
    $requiredGraphModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Users", "Microsoft.Graph.Sites", "Microsoft.Graph.Groups")
    
    foreach ($module in $requiredGraphModules) {
        if (!(Get-Module -ListAvailable -Name $module)) {
            Write-Log "Installing required Microsoft Graph module: $module" "WARN"
            Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser
        }
    }
    
    # Import required modules
    Write-Log "Importing Microsoft Graph modules..."
    Import-Module Microsoft.Graph.Authentication -Force
    Import-Module Microsoft.Graph.Users -Force
    Import-Module Microsoft.Graph.Sites -Force
    Import-Module Microsoft.Graph.Groups -Force
    
    # Load existing CSV data
    $existingData = Get-ExistingCsvData -CsvPath $CsvFilePath
    
    # Connect to Microsoft Graph
    Connect-ToMicrosoftGraph
    
    # Process users based on mode
    if ($useHardCodedUsers) {
        # Process hard-coded user list
        $newResults = Get-OneDriveUrlsFromUserList -UserList $UserList -OutputPath $CsvFilePath
        $sourceDescription = "Hard-coded user list ($($UserList.Count) users)"
    } else {
        # Process multiple groups query
        $newResults = Get-OneDriveUrlsFromGroups -GroupNames $GroupNames -OutputPath $CsvFilePath
        $sourceDescription = "Entra ID Groups: $($GroupNames -join ', ') (Total: $($newResults.Count) unique users)"
    }
    
    # Merge with existing data
    $finalData = Merge-CsvData -ExistingData $existingData -NewResults $newResults
    
    # Export merged results to CSV
    Write-Log "Exporting merged results to: $CsvFilePath"
    $finalData | Export-Csv -Path $CsvFilePath -NoTypeInformation
    
    # Display summary
    $successCount = ($newResults | Where-Object { $_.Status -eq "Success" }).Count
    $errorCount = ($newResults | Where-Object { $_.Status -eq "Error" }).Count
    $noOneDriveCount = ($newResults | Where-Object { $_.Status -eq "No OneDrive" }).Count
    
    Write-Log "=== PROCESSING SUMMARY ===" "INFO"
    Write-Log "Source: $sourceDescription"
    Write-Log "Total users processed: $($newResults.Count)"
    Write-Log "Successful: $successCount"
    Write-Log "No OneDrive: $noOneDriveCount"
    Write-Log "Errors: $errorCount"
    Write-Log "Total entries in CSV: $($finalData.Count)"
    Write-Log "Results saved to: $CsvFilePath"
    
    # Run cleanup script if requested
    if ($RunCleanupAfter) {
        Write-Log "=== LAUNCHING CLEANUP SCRIPT ===" "INFO"
        
        if ([string]::IsNullOrEmpty($CleanupScriptPath)) {
            # Assume cleanup script is in same directory
            $CleanupScriptPath = Join-Path (Split-Path $MyInvocation.MyCommand.Path) "OneDrive-Cleanup.ps1"
        }
        
        if (!(Test-Path $CleanupScriptPath)) {
            Write-Log "Cleanup script not found: $CleanupScriptPath" "ERROR"
            exit 1
        }
        
        if ([string]::IsNullOrEmpty($SharePointAdminUrl)) {
            Write-Log "SharePointAdminUrl is required when RunCleanupAfter is specified" "ERROR"
            Write-Log "Please set SharePointAdminUrl parameter or update the configuration section in the script" "ERROR"
            exit 1
        }
        
        # Build parameters for cleanup script
        $cleanupParams = @{
            CsvFilePath = $CsvFilePath
            SharePointAdminUrl = $SharePointAdminUrl
        }
        
        # Add authentication parameters
        if ($UseInteractiveAuth) {
            $cleanupParams.UseInteractiveAuth = $true
            $cleanupParams.InteractiveClientId = $ClientId
            $cleanupParams.TenantId = $TenantId
            if ($Username) { $cleanupParams.Username = $Username }
            if ($Password) { $cleanupParams.Password = $Password }
        } else {
            $cleanupParams.TenantId = $TenantId
            $cleanupParams.ClientId = $ClientId
            $cleanupParams.Thumbprint = $Thumbprint
        }
        
        if ($WhatIfCleanup) {
            $cleanupParams.WhatIf = $true
        }
        
        Write-Log "Starting cleanup script: $CleanupScriptPath"
        Write-Log "Cleanup parameters: CSV=$CsvFilePath, SharePoint=$SharePointAdminUrl, WhatIf=$WhatIfCleanup"
        
        # Execute cleanup script
        & $CleanupScriptPath @cleanupParams
    }
    
}
catch {
    Write-Log "Critical error: $($_.Exception.Message)" "ERROR"
    throw
}
finally {
    # Disconnect from Microsoft Graph
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-Log "Disconnected from Microsoft Graph"
    }
    catch {
        # Ignore disconnect errors
    }
}

<#
.SYNOPSIS
Queries an Entra ID group and updates CSV file with member OneDrive URLs.

.DESCRIPTION
This script queries an Entra ID group to get member OneDrive URLs and updates an existing CSV file or creates a new one.
It merges new results with existing data, preserving entries not in the current group query.
Optionally can launch the cleanup script after completion.
Log files are automatically organized into a "GroupQueryLogs" folder.

.PARAMETER GroupNames
Array of Entra ID group names to query for OneDrive URLs (supports multiple groups)

.PARAMETER GroupName
Single Entra ID group name to query (alternative to GroupNames for single group)

.PARAMETER UserList
Array of User Principal Names to process instead of group query (overrides configured HardCodedUsers)

.PARAMETER TenantId
Your Azure AD Tenant ID (defaults to configured value)

.PARAMETER ClientId
The Client ID of your Azure App Registration (defaults to configured value)

.PARAMETER Thumbprint
The certificate thumbprint for certificate authentication (defaults to configured value, required if not using interactive auth)

.PARAMETER UseInteractiveAuth
Use interactive authentication instead of certificate authentication

.PARAMETER Username
Username for interactive authentication (optional)

.PARAMETER Password
SecureString password for interactive authentication (optional)

.PARAMETER CsvFilePath
Path to CSV file to update/create (defaults to OneDriveURLs.csv)

.PARAMETER LogFilePath
Path for log file (optional, defaults to timestamped file in GroupQueryLogs folder)

.PARAMETER CleanupScriptPath
Path to the cleanup script (optional, defaults to OneDrive-Cleanup.ps1 in same directory)

.PARAMETER SharePointAdminUrl
SharePoint Admin Center URL (defaults to configured value, required if RunCleanupAfter is true)

.PARAMETER RunCleanupAfter
Launch the cleanup script after group query completes (defaults to true)

.PARAMETER WhatIfCleanup
Pass WhatIf parameter to cleanup script

.EXAMPLE
# Use configured default group name (certificate auth)
.\GroupQuery.ps1

.EXAMPLE
# Use configured default group name with interactive auth and run cleanup
.\GroupQuery.ps1 -UseInteractiveAuth -RunCleanupAfter -WhatIfCleanup

.EXAMPLE
# Override configured group name
.\GroupQuery.ps1 -GroupName "Different OneDrive Group"

.EXAMPLE
# Override configured values
.\GroupQuery.ps1 -GroupNames @("Different Group 1","Different Group 2") -TenantId "different-tenant-id"

.EXAMPLE
# Update specific CSV file
.\GroupQuery.ps1 -CsvFilePath "C:\Scripts\MyOneDriveList.csv"

.NOTES
CONFIGURATION: Update the configuration section at the top of this script with your tenant values:
- TenantId: Your Azure AD Tenant ID
- ClientId: Your App Registration Client ID  
- Thumbprint: Your certificate thumbprint
- SharePointAdminUrl: Your SharePoint Admin Center URL
- DefaultGroupNames: Array of default group names to query (optional)
- HardCodedUsers: Array of specific users to process (optional, overrides group query)

LOG FILES: 
- Log files are automatically organized into a "GroupQueryLogs" folder
- The folder is created automatically if it doesn't exist
- Default log file name format: GroupQuery_YYYYMMDD_HHMMSS.log

Multiple Groups Support:
- Processes all specified groups and combines results
- Automatically removes duplicate users across groups
- Adds SourceGroup column to track which group each user came from

Usage Priority (in order of precedence):
1. Hard-coded users in configuration (if configured) 
2. GroupNames parameter (if provided)
3. GroupName parameter (if provided, converted to GroupNames array)
4. DefaultGroupNames in configuration (if configured)

Required Permissions for App Registration:
- Microsoft Graph API:
  * User.Read.All (to read user information)
  * Sites.Read.All (to read OneDrive site information)  
  * Group.Read.All (to read group information)
  * GroupMember.Read.All (to read group membership)

CSV File Format:
The script maintains these columns:
- UserPrincipalName
- DisplayName  
- OneDriveURL
- Status
- ErrorMessage
- LastUpdated

Module Requirements:
- Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Sites, Microsoft.Graph.Groups (auto-installed)
#>
