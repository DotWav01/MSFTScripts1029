<#
.SYNOPSIS
    Updates retention policy with OneDrive URLs from Entra ID groups or CSV file
.DESCRIPTION
    Can query Entra ID groups to get member OneDrive URLs, or read from CSV file.
    Tracks which URLs have been added to the retention policy.
    Supports both interactive user authentication and app-only certificate authentication.
.PARAMETER ConfigPath
    Path to the configuration JSON filez
.PARAMETER AuthMethod
    Override authentication method from config. Values: 'Interactive' or 'Certificate'
.PARAMETER QueryGroups
    Switch to query Entra ID groups instead of using CSV file
.PARAMETER GroupNames
    Array of Entra ID group names to query (only used with -QueryGroups)
#>

param(
    [string]$ConfigPath = "",
    [ValidateSet('Interactive', 'Certificate', '')]
    [string]$AuthMethod = "",
    [switch]$QueryGroups,
    [string[]]$GroupNames
)

# Function to write log messages
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO",
        [string]$LogPath = $null
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Console output with colors
    switch ($Level) {
        "ERROR" { Write-Host $logMessage -ForegroundColor Red }
        "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        default { Write-Host $logMessage -ForegroundColor Cyan }
    }
    
    # Write to log file if path provided
    if ($LogPath) {
        try {
            $logDir = Split-Path -Path $LogPath -Parent
            if (-not (Test-Path $logDir)) {
                New-Item -Path $logDir -ItemType Directory -Force | Out-Null
            }
            Add-Content -Path $LogPath -Value $logMessage
        }
        catch {
            Write-Host "Warning: Could not write to log file: $_" -ForegroundColor Yellow
        }
    }
}

# Function to connect to Microsoft Graph
function Connect-ToMicrosoftGraph {
    param(
        [string]$TenantId,
        [string]$AppId,
        [string]$CertificateThumbprint,
        [string]$CertificatePath,
        [string]$CertificatePassword,
        [bool]$UseInteractive,
        [string]$LogPath
    )
    
    Write-Log "Connecting to Microsoft Graph..." "INFO" $LogPath
    
    # Check if Microsoft Graph modules are installed
    $requiredGraphModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Users", "Microsoft.Graph.Groups")
    
    foreach ($module in $requiredGraphModules) {
        if (!(Get-Module -ListAvailable -Name $module)) {
            Write-Log "Installing required Microsoft Graph module: $module" "WARNING" $LogPath
            Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser
        }
    }
    
    # Import required modules
    Import-Module Microsoft.Graph.Authentication -Force
    Import-Module Microsoft.Graph.Users -Force
    Import-Module Microsoft.Graph.Groups -Force
    
    try {
        if ($UseInteractive) {
            # Interactive authentication
            Write-Log "Using interactive authentication for Microsoft Graph..." "INFO" $LogPath
            if ($AppId) {
                Connect-MgGraph -ClientId $AppId -TenantId $TenantId -Scopes "User.Read.All", "Group.Read.All" -ErrorAction Stop
            } else {
                Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All" -ErrorAction Stop
            }
        }
        else {
            # Certificate authentication
            if ($CertificatePath -and $CertificatePassword) {
                Write-Log "Using certificate from file for Microsoft Graph..." "INFO" $LogPath
                $certPassword = ConvertTo-SecureString -String $CertificatePassword -AsPlainText -Force
                $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(
                    $CertificatePath, 
                    $certPassword
                )
                Connect-MgGraph -ClientId $AppId -TenantId $TenantId -Certificate $cert -ErrorAction Stop
            }
            elseif ($CertificateThumbprint) {
                Write-Log "Using certificate from store for Microsoft Graph..." "INFO" $LogPath
                Connect-MgGraph -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint -ErrorAction Stop
            }
            else {
                throw "Certificate authentication requires either certificatePath+certificatePassword OR certificateThumbprint"
            }
        }
        
        Write-Log "Successfully connected to Microsoft Graph" "SUCCESS" $LogPath
    }
    catch {
        Write-Log "Failed to connect to Microsoft Graph: $_" "ERROR" $LogPath
        throw
    }
}

# Function to get OneDrive URLs from Entra ID groups
function Get-OneDriveUrlsFromGroups {
    param(
        [string[]]$GroupNames,
        [string]$LogPath
    )
    
    Write-Log "=== QUERYING ENTRA ID GROUPS FOR ONEDRIVE URLS ===" "INFO" $LogPath
    Write-Log "Groups to process: $($GroupNames -join ', ')" "INFO" $LogPath
    
    $allResults = @()
    $groupCounter = 0
    
    foreach ($groupName in $GroupNames) {
        $groupCounter++
        Write-Log "Processing group $groupCounter of $($GroupNames.Count): $groupName" "INFO" $LogPath
        
        try {
            # Find the group by name
            Write-Log "  Searching for group: $groupName" "INFO" $LogPath
            $group = Get-MgGroup -Filter "displayName eq '$groupName'" -ErrorAction Stop
            
            if (!$group) {
                Write-Log "  Group '$groupName' not found" "WARNING" $LogPath
                continue
            }
            
            if ($group.Count -gt 1) {
                Write-Log "  Multiple groups found with name '$groupName'. Using the first one." "WARNING" $LogPath
                $group = $group[0]
            }
            
            Write-Log "  Found group: $($group.DisplayName) (ID: $($group.Id))" "SUCCESS" $LogPath
            
            # Get group members
            Write-Log "  Retrieving group members..." "INFO" $LogPath
            $groupMembers = Get-MgGroupMember -GroupId $group.Id -All
            
            # Filter to only get user members
            $users = $groupMembers | Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.user' }
            
            if ($users.Count -eq 0) {
                Write-Log "  No user members found in group '$groupName'" "WARNING" $LogPath
                continue
            }
            
            Write-Log "  Found $($users.Count) user members in the group" "INFO" $LogPath
            
            # Process users in this group
            $userCounter = 0
            foreach ($userMember in $users) {
                $userCounter++
                $userId = $userMember.Id
                
                Write-Log "  Processing user $userCounter of $($users.Count) from group '$groupName'..." "INFO" $LogPath
                
                try {
                    # Get full user information
                    $mgUser = Get-MgUser -UserId $userId -ErrorAction Stop
                    $userPrincipalName = $mgUser.UserPrincipalName
                    
                    Write-Log "    User: $userPrincipalName" "INFO" $LogPath
                    
                    # Check if user already processed (avoid duplicates across groups)
                    $existingUser = $allResults | Where-Object { $_.UserPrincipalName -eq $userPrincipalName }
                    if ($existingUser) {
                        Write-Log "    User already processed from another group, skipping duplicate" "INFO" $LogPath
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
                        Write-Log "    OneDrive URL found: $oneDriveUrl" "SUCCESS" $LogPath
                    } else {
                        $oneDriveUrl = ""
                        Write-Log "    OneDrive not found" "WARNING" $LogPath
                    }
                }
                catch {
                    $userPrincipalName = if ($mgUser) { $mgUser.UserPrincipalName } else { "Unknown" }
                    $oneDriveUrl = ""
                    Write-Log "    Error: $($_.Exception.Message)" "ERROR" $LogPath
                }
                
                # Only add if OneDrive URL was found
                if ($oneDriveUrl) {
                    $allResults += [PSCustomObject]@{
                        UserPrincipalName = $userPrincipalName
                        DisplayName = if ($mgUser) { $mgUser.DisplayName } else { "N/A" }
                        OneDriveURL = $oneDriveUrl
                        SourceGroup = $groupName
                        AddedToPolicy = ""
                    }
                }
                
                # Add small delay to avoid throttling
                Start-Sleep -Milliseconds 200
            }
            
            Write-Log "  Completed processing group '$groupName'" "SUCCESS" $LogPath
        }
        catch {
            Write-Log "  Error processing group '$groupName': $($_.Exception.Message)" "ERROR" $LogPath
        }
    }
    
    Write-Log "Total unique users with OneDrive found: $($allResults.Count)" "INFO" $LogPath
    return $allResults
}

# Load configuration
Write-Log "Loading configuration from $ConfigPath"
try {
    if (-not (Test-Path $ConfigPath)) {
        throw "Configuration file not found at: $ConfigPath"
    }
    
    $config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
    Write-Log "Configuration loaded successfully" "SUCCESS"
}
catch {
    Write-Log "Failed to load configuration: $_" "ERROR"
    exit 1
}

# Set log path from config
$logPath = if ($config.logPath) { $config.logPath } else { $null }

# Determine authentication method (parameter overrides config)
$selectedAuthMethod = if ($AuthMethod) { 
    $AuthMethod 
} elseif ($config.authMethod) { 
    $config.authMethod 
} else { 
    "Interactive"  # Default to Interactive if not specified
}

Write-Log "Authentication method: $selectedAuthMethod" "INFO" $logPath

# Validate required configuration values
$requiredFields = @('policyName', 'csvPath')
foreach ($field in $requiredFields) {
    if ([string]::IsNullOrWhiteSpace($config.$field)) {
        Write-Log "Missing required configuration field: $field" "ERROR" $logPath
        exit 1
    }
}

# If QueryGroups switch is used, validate group configuration
if ($QueryGroups) {
    if ($GroupNames -and $GroupNames.Count -gt 0) {
        # Use provided group names
        Write-Log "Using provided group names: $($GroupNames -join ', ')" "INFO" $logPath
    }
    elseif ($config.groupNames -and $config.groupNames.Count -gt 0) {
        # Use group names from config
        $GroupNames = $config.groupNames
        Write-Log "Using group names from config: $($GroupNames -join ', ')" "INFO" $logPath
    }
    else {
        Write-Log "QueryGroups switch used but no group names provided in parameter or config" "ERROR" $logPath
        exit 1
    }
    
    # Validate we have tenant/app info for Graph connection
    if ([string]::IsNullOrWhiteSpace($config.tenantId) -or [string]::IsNullOrWhiteSpace($config.appId)) {
        Write-Log "QueryGroups requires 'tenantId' and 'appId' in config for Microsoft Graph connection" "ERROR" $logPath
        exit 1
    }
}

# Validate certificate configuration if using certificate auth
if ($selectedAuthMethod -eq "Certificate") {
    if ([string]::IsNullOrWhiteSpace($config.appId) -or [string]::IsNullOrWhiteSpace($config.organization)) {
        Write-Log "Certificate authentication requires 'appId' and 'organization' in config" "ERROR" $logPath
        exit 1
    }
    
    $useCertFile = $false
    if ($config.certificatePath -and $config.certificatePassword) {
        $useCertFile = $true
        if (-not (Test-Path $config.certificatePath)) {
            Write-Log "Certificate file not found at: $($config.certificatePath)" "ERROR" $logPath
            exit 1
        }
    }
    elseif ($config.certificateThumbprint) {
        $useCertFile = $false
    }
    else {
        Write-Log "Certificate authentication requires either certificatePath+certificatePassword OR certificateThumbprint" "ERROR" $logPath
        exit 1
    }
}

# Set batch size (default to 100 if not specified)
$batchSize = if ($config.batchSize) { $config.batchSize } else { 100 }

# If QueryGroups is enabled, query Entra ID groups first
if ($QueryGroups) {
    Write-Log "=== GROUP QUERY MODE ENABLED ===" "INFO" $logPath
    
    # Connect to Microsoft Graph
    Connect-ToMicrosoftGraph -TenantId $config.tenantId `
                             -AppId $config.appId `
                             -CertificateThumbprint $config.certificateThumbprint `
                             -CertificatePath $config.certificatePath `
                             -CertificatePassword $config.certificatePassword `
                             -UseInteractive ($selectedAuthMethod -eq "Interactive") `
                             -LogPath $logPath
    
    # Query groups and get OneDrive URLs
    $groupResults = Get-OneDriveUrlsFromGroups -GroupNames $GroupNames -LogPath $logPath
    
    if ($groupResults.Count -eq 0) {
        Write-Log "No OneDrive URLs found from group query" "WARNING" $logPath
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        exit 0
    }
    
    # Load existing CSV if it exists
    if (Test-Path $config.csvPath) {
        Write-Log "Loading existing CSV to merge with group results..." "INFO" $logPath
        $existingCsv = Import-Csv -Path $config.csvPath
        
        # Merge results - update existing entries and add new ones
        $existingHash = @{}
        foreach ($item in $existingCsv) {
            if ($item.OneDriveURL) {
                $existingHash[$item.OneDriveURL] = $item
            }
        }
        
        foreach ($groupItem in $groupResults) {
            if ($existingHash.ContainsKey($groupItem.OneDriveURL)) {
                # Update existing entry
                $existingHash[$groupItem.OneDriveURL].UserPrincipalName = $groupItem.UserPrincipalName
                $existingHash[$groupItem.OneDriveURL].DisplayName = $groupItem.DisplayName
                if ($groupItem.SourceGroup) {
                    $existingHash[$groupItem.OneDriveURL] | Add-Member -NotePropertyName "SourceGroup" -NotePropertyValue $groupItem.SourceGroup -Force
                }
            }
            else {
                # Add new entry
                $existingHash[$groupItem.OneDriveURL] = $groupItem
            }
        }
        
        $oneDriveList = $existingHash.Values | ForEach-Object { $_ }
    }
    else {
        # No existing CSV, use group results directly
        $oneDriveList = $groupResults
    }
    
    # Save updated CSV
    Write-Log "Saving updated CSV with group results..." "INFO" $logPath
    $oneDriveList | Export-Csv -Path $config.csvPath -NoTypeInformation -Force
    Write-Log "CSV updated with $($oneDriveList.Count) total entries" "SUCCESS" $logPath
    
    # Disconnect from Microsoft Graph
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Write-Log "Disconnected from Microsoft Graph" "INFO" $logPath
}

# Connect to Security & Compliance Center
Write-Log "Connecting to Security & Compliance Center..." "INFO" $logPath
try {
    if ($selectedAuthMethod -eq "Certificate") {
        # Certificate-based authentication (App-only)
        if ($useCertFile) {
            # Load certificate from PFX file
            Write-Log "Using certificate from file: $($config.certificatePath)" "INFO" $logPath
            $certPassword = ConvertTo-SecureString -String $config.certificatePassword -AsPlainText -Force
            $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(
                $config.certificatePath, 
                $certPassword
            )
            
            Connect-IPPSSession -AppId $config.appId `
                                -Certificate $cert `
                                -Organization $config.organization
        }
        else {
            # Use certificate from store via thumbprint
            Write-Log "Using certificate from store with thumbprint: $($config.certificateThumbprint)" "INFO" $logPath
            Connect-IPPSSession -AppId $config.appId `
                                -CertificateThumbprint $config.certificateThumbprint `
                                -Organization $config.organization
        }
        Write-Log "Connected successfully using certificate authentication" "SUCCESS" $logPath
    }
    else {
        # Interactive authentication (User account)
        Write-Log "Using interactive authentication - login prompt will appear..." "INFO" $logPath
        Connect-IPPSSession
        Write-Log "Connected successfully using interactive authentication" "SUCCESS" $logPath
    }
}
catch {
    Write-Log "Connection failed: $_" "ERROR" $logPath
    if ($selectedAuthMethod -eq "Certificate") {
        Write-Log "Ensure the certificate is valid and the app has proper permissions" "ERROR" $logPath
    }
    else {
        Write-Log "Ensure you have proper permissions to access Security & Compliance Center" "ERROR" $logPath
    }
    exit 1
}

# Verify CSV file exists
if (-not (Test-Path $config.csvPath)) {
    Write-Log "CSV file not found at: $($config.csvPath)" "ERROR" $logPath
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    exit 1
}

# Import CSV and ensure AddedToPolicy column exists
Write-Log "Loading CSV file from: $($config.csvPath)" "INFO" $logPath
try {
    $oneDriveList = Import-Csv -Path $config.csvPath
    
    if ($oneDriveList.Count -eq 0) {
        Write-Log "CSV file is empty" "WARNING" $logPath
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        exit 0
    }
    
    # Validate CSV structure
    if (-not ($oneDriveList[0].PSObject.Properties.Name -contains "OneDriveURL")) {
        Write-Log "CSV must contain 'OneDriveURL' column" "ERROR" $logPath
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        exit 1
    }
    
    # Add AddedToPolicy column if it doesn't exist
    if (-not ($oneDriveList[0].PSObject.Properties.Name -contains "AddedToPolicy")) {
        Write-Log "Adding 'AddedToPolicy' column to CSV entries" "INFO" $logPath
        $oneDriveList | ForEach-Object {
            $_ | Add-Member -MemberType NoteProperty -Name "AddedToPolicy" -Value "" -Force
        }
    }
    
    Write-Log "CSV loaded successfully - Total entries: $($oneDriveList.Count)" "SUCCESS" $logPath
}
catch {
    Write-Log "Failed to load CSV: $_" "ERROR" $logPath
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    exit 1
}

# Get current policy locations
Write-Log "Retrieving current retention policy: $($config.policyName)" "INFO" $logPath
try {
    $policy = Get-RetentionCompliancePolicy -Identity $config.policyName -DistributionDetail
    $existingLocations = $policy.OneDriveLocation
    Write-Log "Found $($existingLocations.Count) existing locations in policy" "SUCCESS" $logPath
}
catch {
    Write-Log "Error retrieving policy: $_" "ERROR" $logPath
    Write-Log "Please verify the policy name and your permissions" "ERROR" $logPath
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    exit 1
}

# Filter URLs to add
Write-Log "Processing URLs..." "INFO" $logPath
$urlsToAdd = @()
$skippedCount = 0
$urlToItemMap = @{}

foreach ($item in $oneDriveList) {
    # Skip if no URL provided
    if ([string]::IsNullOrWhiteSpace($item.OneDriveURL)) {
        Write-Log "Skipping empty URL entry" "WARNING" $logPath
        continue
    }
    
    # Skip if already marked as added in CSV
    if ($item.AddedToPolicy -eq "Yes") {
        $skippedCount++
        continue
    }
    
    # Skip if already exists in policy
    if ($existingLocations -contains $item.OneDriveURL) {
        $item.AddedToPolicy = "Yes"
        $skippedCount++
        continue
    }
    
    # Add to processing list
    $urlsToAdd += $item.OneDriveURL
    $urlToItemMap[$item.OneDriveURL] = $item
}

Write-Log "URLs to add: $($urlsToAdd.Count) | Already added/skipped: $skippedCount" "INFO" $logPath

# Process URLs in batches
if ($urlsToAdd.Count -gt 0) {
    $totalBatches = [Math]::Ceiling($urlsToAdd.Count / $batchSize)
    $successfullyAdded = @()
    $failedUrls = @()

    Write-Log "Starting batch processing ($totalBatches batches of max $batchSize URLs each)..." "INFO" $logPath

    for ($i = 0; $i -lt $urlsToAdd.Count; $i += $batchSize) {
        $currentBatch = [Math]::Floor($i / $batchSize) + 1
        $endIndex = [Math]::Min($i + $batchSize - 1, $urlsToAdd.Count - 1)
        $batch = $urlsToAdd[$i..$endIndex]
        
        Write-Log "Processing batch $currentBatch of $totalBatches ($($batch.Count) URLs)..." "INFO" $logPath
        
        try {
            Set-RetentionCompliancePolicy -Identity $config.policyName -AddOneDriveLocation $batch -ErrorAction Stop
            Write-Log "Batch $currentBatch completed successfully" "SUCCESS" $logPath
            $successfullyAdded += $batch
        }
        catch {
            Write-Log "Batch $currentBatch failed: $_" "ERROR" $logPath
            $failedUrls += $batch
        }
        
        # Throttle between batches
        if ($currentBatch -lt $totalBatches) {
            Start-Sleep -Seconds 2
        }
    }

    # Update CSV with results
    Write-Log "Updating CSV file..." "INFO" $logPath
    foreach ($url in $successfullyAdded) {
        if ($urlToItemMap.ContainsKey($url)) {
            $urlToItemMap[$url].AddedToPolicy = "Yes"
        }
    }

    # Save updated CSV
    try {
        $oneDriveList | Export-Csv -Path $config.csvPath -NoTypeInformation -Force
        Write-Log "CSV file updated successfully" "SUCCESS" $logPath
    }
    catch {
        Write-Log "Failed to update CSV file: $_" "ERROR" $logPath
    }

    # Summary
    Write-Log "================================================" "INFO" $logPath
    Write-Log "SUMMARY" "INFO" $logPath
    Write-Log "================================================" "INFO" $logPath
    Write-Log "Authentication Method: $selectedAuthMethod" "INFO" $logPath
    if ($QueryGroups) {
        Write-Log "Source: Entra ID Groups ($($GroupNames -join ', '))" "INFO" $logPath
    }
    Write-Log "Total URLs processed: $($urlsToAdd.Count)" "INFO" $logPath
    Write-Log "Successfully added: $($successfullyAdded.Count)" "SUCCESS" $logPath
    Write-Log "Failed: $($failedUrls.Count)" $(if ($failedUrls.Count -gt 0) { "ERROR" } else { "INFO" }) $logPath
    Write-Log "Skipped (already added): $skippedCount" "INFO" $logPath
    Write-Log "================================================" "INFO" $logPath
    
    if ($failedUrls.Count -gt 0) {
        Write-Log "Failed URLs:" "ERROR" $logPath
        $failedUrls | ForEach-Object { Write-Log "  - $_" "ERROR" $logPath }
    }
}
else {
    Write-Log "No new URLs to add - all entries already in policy" "SUCCESS" $logPath
}

# Disconnect
Write-Log "Disconnecting from Security & Compliance Center..." "INFO" $logPath
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
Write-Log "Script completed" "SUCCESS" $logPath
