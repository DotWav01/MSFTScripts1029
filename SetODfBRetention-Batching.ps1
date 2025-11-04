<#
.SYNOPSIS
    Updates retention policy with OneDrive URLs from Entra ID groups or CSV file with multi-policy support
.DESCRIPTION
    Can query Entra ID groups to get member OneDrive URLs, or read from CSV file.
    Tracks which URLs have been added to retention policies.
    Uses smart batch processing to efficiently handle multiple policies with 100 URL capacity limits.
    Supports both interactive user authentication and app-only certificate authentication.
.PARAMETER ConfigPath
    Path to the configuration JSON file
.PARAMETER AuthMethod
    Override authentication method from config. Values: 'Interactive' or 'Certificate'
.PARAMETER QueryGroups
    Switch to query Entra ID groups instead of using CSV file
.PARAMETER GroupNames
    Array of Entra ID group names to query (only used with -QueryGroups)
.PARAMETER AutoConfirm
    Skip confirmation prompt and proceed automatically (useful for scheduled tasks)
#>

param(
    [string]$ConfigPath = "C:\Secure\config.json",
    [ValidateSet('Interactive', 'Certificate', '')]
    [string]$AuthMethod = "",
    [switch]$QueryGroups,
    [string[]]$GroupNames,
    [switch]$AutoConfirm
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
                        PolicyName = ""
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
$requiredFields = @('csvPath')
foreach ($field in $requiredFields) {
    if ([string]::IsNullOrWhiteSpace($config.$field)) {
        Write-Log "Missing required configuration field: $field" "ERROR" $logPath
        exit 1
    }
}

# Validate policy configuration - support both single and multiple policies
if ($config.policyName) {
    # Single policy (backward compatibility)
    $policyNames = @($config.policyName)
} elseif ($config.policyNames -and $config.policyNames.Count -gt 0) {
    # Multiple policies
    $policyNames = $config.policyNames
} else {
    Write-Log "Missing required configuration: either 'policyName' or 'policyNames' must be specified" "ERROR" $logPath
    exit 1
}

Write-Log "Retention policies configured: $($policyNames -join ', ')" "INFO" $logPath

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
    
    # Load existing CSV if it exists and merge
    if (Test-Path $config.csvPath) {
        Write-Log "Loading existing CSV to merge with group results..." "INFO" $logPath
        $existingCsv = Import-Csv -Path $config.csvPath
        
        # Get list of existing OneDrive URLs
        $existingUrls = $existingCsv | Select-Object -ExpandProperty OneDriveURL
        
        # Filter group results to only new URLs
        $newEntries = $groupResults | Where-Object { $_.OneDriveURL -notin $existingUrls }
        
        # Combine existing and new
        $oneDriveList = @($existingCsv) + @($newEntries)
        
        Write-Log "Found $($newEntries.Count) new entries from groups to add to existing $($existingCsv.Count) entries" "INFO" $logPath
    }
    else {
        # No existing CSV, use group results directly
        $oneDriveList = $groupResults
        Write-Log "No existing CSV found, using group results only" "INFO" $logPath
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

# Import CSV and ensure required columns exist
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
    
    # Add required columns if they don't exist
    if (-not ($oneDriveList[0].PSObject.Properties.Name -contains "AddedToPolicy")) {
        Write-Log "Adding 'AddedToPolicy' column to CSV entries" "INFO" $logPath
        $oneDriveList | ForEach-Object {
            $_ | Add-Member -MemberType NoteProperty -Name "AddedToPolicy" -Value "" -Force
        }
    }
    
    if (-not ($oneDriveList[0].PSObject.Properties.Name -contains "PolicyName")) {
        Write-Log "Adding 'PolicyName' column to CSV entries" "INFO" $logPath
        $oneDriveList | ForEach-Object {
            $_ | Add-Member -MemberType NoteProperty -Name "PolicyName" -Value "" -Force
        }
    }
    
    Write-Log "CSV loaded successfully - Total entries: $($oneDriveList.Count)" "SUCCESS" $logPath
}
catch {
    Write-Log "Failed to load CSV: $_" "ERROR" $logPath
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    exit 1
}

# Get current locations for all policies
Write-Log "Retrieving current retention policies..." "INFO" $logPath
$policyLocations = @{}
foreach ($policyName in $policyNames) {
    try {
        Write-Log "  Checking policy: $policyName" "INFO" $logPath
        $policy = Get-RetentionCompliancePolicy -Identity $policyName -DistributionDetail -ErrorAction Stop
        $policyLocations[$policyName] = $policy.OneDriveLocation
        Write-Log "  Found $($policy.OneDriveLocation.Count) existing locations in policy '$policyName'" "SUCCESS" $logPath
    }
    catch {
        Write-Log "  Error retrieving policy '$policyName': $_" "ERROR" $logPath
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        exit 1
    }
}

# Process URLs with smart batching by policy
Write-Log "=== PROCESSING ONEDRIVE URLS WITH SMART POLICY BATCHING ===" "INFO" $logPath
$urlsToProcess = @()
$skippedCount = 0
$urlToItemMap = @{}

# First pass: identify URLs that need to be added
foreach ($item in $oneDriveList) {
    # Skip if no URL provided
    if ([string]::IsNullOrWhiteSpace($item.OneDriveURL)) {
        Write-Log "Skipping empty URL entry" "WARNING" $logPath
        continue
    }
    
    # Check if already added to any policy
    $alreadyAdded = $false
    foreach ($policyName in $policyNames) {
        if ($policyLocations[$policyName] -contains $item.OneDriveURL) {
            if ($item.AddedToPolicy -ne "Yes") {
                $item.AddedToPolicy = "Yes"
                $item.PolicyName = $policyName
            }
            $alreadyAdded = $true
            $skippedCount++
            break
        }
    }
    
    # Also skip if marked as added in CSV
    if (!$alreadyAdded -and $item.AddedToPolicy -eq "Yes") {
        $skippedCount++
        $alreadyAdded = $true
    }
    
    if (!$alreadyAdded) {
        # Add to processing list
        $urlsToProcess += $item.OneDriveURL
        $urlToItemMap[$item.OneDriveURL] = $item
    }
}

Write-Log "URLs to process: $($urlsToProcess.Count) | Already added/skipped: $skippedCount" "INFO" $logPath

# Process URLs with smart batching by policy
if ($urlsToProcess.Count -gt 0) {
    $successfullyAdded = @()
    $failedUrls = @()
    $policyCounts = @{}
    
    # Initialize policy counters
    foreach ($policyName in $policyNames) {
        $policyCounts[$policyName] = 0
    }
    
    Write-Log "Using batch processing strategy for efficient policy updates" "INFO" $logPath
    
    $remainingUrls = [System.Collections.ArrayList]@($urlsToProcess)
    
    # Build batches for each policy
    $policyBatches = @{}
    
    Write-Log "" "INFO" $logPath
    Write-Log "========================================" "INFO" $logPath
    Write-Log "ANALYZING POLICIES AND BUILDING BATCHES" "INFO" $logPath
    Write-Log "========================================" "INFO" $logPath
    
    foreach ($policyName in $policyNames) {
        if ($remainingUrls.Count -eq 0) {
            Write-Log "All URLs have been assigned to policies" "SUCCESS" $logPath
            break
        }
        
        Write-Log "" "INFO" $logPath
        Write-Log "Analyzing policy: $policyName" "INFO" $logPath
        
        # Check current capacity of the policy
        try {
            $policy = Get-RetentionCompliancePolicy -Identity $policyName -DistributionDetail -ErrorAction Stop
            $currentCount = $policy.OneDriveLocation.Count
            $availableCapacity = 100 - $currentCount
            
            Write-Log "  Current capacity: $currentCount/100 (Available: $availableCapacity slots)" "INFO" $logPath
            
            if ($currentCount -ge 100) {
                Write-Log "  Policy is at full capacity, skipping" "WARNING" $logPath
                continue
            }
            
            # Determine how many URLs we can add to this policy
            $urlsToAddToPolicy = [Math]::Min($availableCapacity, $remainingUrls.Count)
            
            if ($urlsToAddToPolicy -gt 0) {
                # Take URLs from remaining list
                $batchUrls = $remainingUrls[0..($urlsToAddToPolicy - 1)]
                $policyBatches[$policyName] = @{
                    URLs = $batchUrls
                    CurrentCount = $currentCount
                    NewCount = $currentCount + $urlsToAddToPolicy
                }
                
                Write-Log "  Will add $urlsToAddToPolicy URLs to this policy (will become: $($currentCount + $urlsToAddToPolicy)/100)" "SUCCESS" $logPath
                
                # Remove these URLs from remaining list
                foreach ($url in $batchUrls) {
                    $remainingUrls.Remove($url) | Out-Null
                }
            }
        }
        catch {
            Write-Log "  Error checking policy capacity: $_" "ERROR" $logPath
            continue
        }
    }
    
    # Display summary of planned changes
    Write-Log "" "INFO" $logPath
    Write-Log "========================================" "INFO" $logPath
    Write-Log "PLANNED CHANGES SUMMARY" "INFO" $logPath
    Write-Log "========================================" "INFO" $logPath
    Write-Log "Total URLs to be added: $($urlsToProcess.Count - $remainingUrls.Count)" "INFO" $logPath
    Write-Log "URLs that couldn't be assigned (all policies full): $($remainingUrls.Count)" $(if ($remainingUrls.Count -gt 0) { "WARNING" } else { "INFO" }) $logPath
    Write-Log "" "INFO" $logPath
    
    foreach ($policyName in $policyBatches.Keys) {
        $batch = $policyBatches[$policyName]
        Write-Log "Policy: $policyName" "INFO" $logPath
        Write-Log "  Current: $($batch.CurrentCount)/100 → New: $($batch.NewCount)/100" "INFO" $logPath
        Write-Log "  URLs to add: $($batch.URLs.Count)" "INFO" $logPath
        Write-Log "  Users:" "INFO" $logPath
        
        foreach ($url in $batch.URLs) {
            $item = $urlToItemMap[$url]
            $userDisplay = if ($item.UserPrincipalName) { 
                "$($item.UserPrincipalName) ($($item.DisplayName))" 
            } else { 
                $url 
            }
            Write-Log "    - $userDisplay" "INFO" $logPath
        }
        Write-Log "" "INFO" $logPath
    }
    
    # Confirmation prompt
    if ($policyBatches.Count -eq 0) {
        Write-Log "No URLs can be added (all policies are at capacity)" "WARNING" $logPath
    }
    else {
        if (-not $AutoConfirm) {
            Write-Host ""
            Write-Host "========================================" -ForegroundColor Yellow
            Write-Host "CONFIRMATION REQUIRED" -ForegroundColor Yellow
            Write-Host "========================================" -ForegroundColor Yellow
            Write-Host "The script will add $($urlsToProcess.Count - $remainingUrls.Count) URLs across $($policyBatches.Count) retention policies." -ForegroundColor Cyan
            Write-Host ""
            
            $confirmation = Read-Host "Do you want to proceed? (yes/no)"
            
            if ($confirmation -ne "yes" -and $confirmation -ne "y") {
                Write-Log "Operation cancelled by user" "WARNING" $logPath
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
                exit 0
            }
        }
        else {
            Write-Log "AutoConfirm enabled - proceeding without confirmation prompt" "INFO" $logPath
        }
        
        Write-Log "" "INFO" $logPath
        Write-Log "Proceeding with batch additions..." "SUCCESS" $logPath
        
        # Process each policy batch
        foreach ($policyName in $policyBatches.Keys) {
            $batch = $policyBatches[$policyName]
            
            Write-Log "" "INFO" $logPath
            Write-Log "========================================" "INFO" $logPath
            Write-Log "Processing policy: $policyName" "INFO" $logPath
            Write-Log "Adding batch of $($batch.URLs.Count) URLs..." "INFO" $logPath
            Write-Log "========================================" "INFO" $logPath
            
            # Split into sub-batches of 100 if needed (shouldn't happen but just in case)
            $maxBatchSize = 100
            $subBatches = [Math]::Ceiling($batch.URLs.Count / $maxBatchSize)
            
            for ($i = 0; $i -lt $subBatches; $i++) {
                $startIdx = $i * $maxBatchSize
                $endIdx = [Math]::Min($startIdx + $maxBatchSize - 1, $batch.URLs.Count - 1)
                $subBatch = $batch.URLs[$startIdx..$endIdx]
                
                if ($subBatches -gt 1) {
                    Write-Log "Processing sub-batch $($i + 1) of $subBatches ($($subBatch.Count) URLs)..." "INFO" $logPath
                }
                
                $attempt = 0
                $maxAttempts = 3
                $success = $false
                
                while ($attempt -lt $maxAttempts -and -not $success) {
                    $attempt++
                    
                    try {
                        Set-RetentionCompliancePolicy -Identity $policyName -AddOneDriveLocation $subBatch -ErrorAction Stop
                        $success = $true
                        Write-Log "✓ Successfully added batch of $($subBatch.Count) URLs to policy" "SUCCESS" $logPath
                        
                        # Mark all URLs in this batch as added
                        foreach ($url in $subBatch) {
                            $item = $urlToItemMap[$url]
                            $item.AddedToPolicy = "Yes"
                            $item.PolicyName = $policyName
                            $successfullyAdded += $url
                            $policyCounts[$policyName]++
                        }
                    }
                    catch {
                        $errorMessage = $_.Exception.Message
                        
                        if ($errorMessage -like "*being deployed*" -or $errorMessage -like "*being processed*" -or $errorMessage -like "*previous changes*") {
                            if ($attempt -lt $maxAttempts) {
                                Write-Log "Policy is processing previous changes. Waiting 30 seconds before retry (attempt $attempt/$maxAttempts)..." "WARNING" $logPath
                                Start-Sleep -Seconds 30
                            }
                            else {
                                Write-Log "✗ Failed to add batch after $maxAttempts attempts: $errorMessage" "ERROR" $logPath
                                foreach ($url in $subBatch) {
                                    $failedUrls += $url
                                }
                            }
                        }
                        else {
                            Write-Log "✗ Error adding batch to policy: $errorMessage" "ERROR" $logPath
                            foreach ($url in $subBatch) {
                                $failedUrls += $url
                            }
                            break
                        }
                    }
                }
                
                # Small delay between sub-batches if there are multiple
                if ($subBatches -gt 1 -and $i -lt $subBatches - 1) {
                    Start-Sleep -Seconds 5
                }
            }
        }
    }
    
    # Add any remaining URLs to failed list
    foreach ($url in $remainingUrls) {
        if ($url -notin $failedUrls) {
            $failedUrls += $url
        }
    }

    # Save updated CSV
    Write-Log "" "INFO" $logPath
    Write-Log "Updating CSV file..." "INFO" $logPath
    try {
        $oneDriveList | Export-Csv -Path $config.csvPath -NoTypeInformation -Force
        Write-Log "CSV file updated successfully" "SUCCESS" $logPath
    }
    catch {
        Write-Log "Failed to update CSV file: $_" "ERROR" $logPath
    }

    # Summary
    Write-Log "" "INFO" $logPath
    Write-Log "================================================" "INFO" $logPath
    Write-Log "FINAL SUMMARY" "INFO" $logPath
    Write-Log "================================================" "INFO" $logPath
    Write-Log "Authentication Method: $selectedAuthMethod" "INFO" $logPath
    if ($QueryGroups) {
        Write-Log "Source: Entra ID Groups ($($GroupNames -join ', '))" "INFO" $logPath
    }
    Write-Log "Total URLs in CSV: $($oneDriveList.Count)" "INFO" $logPath
    Write-Log "URLs already in policies (skipped): $skippedCount" "INFO" $logPath
    Write-Log "URLs processed in this run: $($urlsToProcess.Count)" "INFO" $logPath
    Write-Log "Successfully added: $($successfullyAdded.Count)" "SUCCESS" $logPath
    Write-Log "Failed: $($failedUrls.Count)" $(if ($failedUrls.Count -gt 0) { "ERROR" } else { "INFO" }) $logPath
    Write-Log "Not added (all policies full): $($remainingUrls.Count)" $(if ($remainingUrls.Count -gt 0) { "WARNING" } else { "INFO" }) $logPath
    Write-Log "" "INFO" $logPath
    Write-Log "Distribution across policies:" "INFO" $logPath
    foreach ($policyName in $policyNames) {
        try {
            $policy = Get-RetentionCompliancePolicy -Identity $policyName -DistributionDetail -ErrorAction Stop
            $currentTotal = $policy.OneDriveLocation.Count
            $status = if ($currentTotal -ge 100) { "FULL" } else { "Available: $(100 - $currentTotal)" }
            Write-Log "  $policyName : $($policyCounts[$policyName]) added in this run | Total: $currentTotal/100 [$status]" "INFO" $logPath
        }
        catch {
            Write-Log "  $policyName : $($policyCounts[$policyName]) added in this run | Unable to verify total" "WARNING" $logPath
        }
    }
    Write-Log "================================================" "INFO" $logPath
    
    if ($failedUrls.Count -gt 0) {
        Write-Log "" "INFO" $logPath
        Write-Log "Failed URLs:" "ERROR" $logPath
        foreach ($url in $failedUrls) {
            $failedItem = $urlToItemMap[$url]
            $userDisplay = if ($failedItem.UserPrincipalName) { $failedItem.UserPrincipalName } else { $url }
            Write-Log "  - $userDisplay" "ERROR" $logPath
        }
    }
    
    if ($remainingUrls.Count -gt 0) {
        Write-Log "" "INFO" $logPath
        Write-Log "URLs not added (all policies at capacity):" "WARNING" $logPath
        foreach ($url in $remainingUrls) {
            $remainingItem = $urlToItemMap[$url]
            $userDisplay = if ($remainingItem.UserPrincipalName) { $remainingItem.UserPrincipalName } else { $url }
            Write-Log "  - $userDisplay" "WARNING" $logPath
        }
        Write-Log "" "WARNING" $logPath
        Write-Log "RECOMMENDATION: Add more retention policies to the configuration to accommodate these users" "WARNING" $logPath
    }
}
else {
    Write-Log "No new URLs to add - all entries already in policies" "SUCCESS" $logPath
}

# Disconnect
Write-Log "" "INFO" $logPath
Write-Log "Disconnecting from Security & Compliance Center..." "INFO" $logPath
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
Write-Log "Script completed" "SUCCESS" $logPath
