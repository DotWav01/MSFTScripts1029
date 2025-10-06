# Windows 365 Cloud PC User Deprovisioning Script
# Author: IT Administrator
# Date: October 2025
# Purpose: Deprovision Windows 365 Cloud PC users by removing them from provisioning and licensing groups

#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Users, Microsoft.Graph.DeviceManagement

<#
.SYNOPSIS
    Deprovisions Windows 365 Cloud PC users by removing them from provisioning and licensing groups.

.DESCRIPTION
    This script removes users from their Windows 365 provisioning groups and licensing groups,
    then checks if their Cloud PCs are in grace period and offers to end the grace period.
    It can process single users or multiple users from a CSV file.

.PARAMETER InputMethod
    Specifies whether to process a single user or multiple users from CSV
    Valid values: "Manual", "CSV"

.PARAMETER UserPrincipalName
    The UPN of the user to deprovision (used with Manual input method)

.PARAMETER CsvPath
    Path to CSV file containing list of users (used with CSV input method)
    CSV should have a column named "UserPrincipalName"

.EXAMPLE
    .\Deprovision-CloudPCUser.ps1 -InputMethod Manual -UserPrincipalName user@domain.com

.EXAMPLE
    .\Deprovision-CloudPCUser.ps1 -InputMethod CSV -CsvPath "C:\Users\list.csv"
#>

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("Manual", "CSV")]
    [string]$InputMethod,
    
    [Parameter(Mandatory=$false)]
    [string]$UserPrincipalName,
    
    [Parameter(Mandatory=$false)]
    [string]$CsvPath
)

# Script variables
$script:LogPath = ".\CloudPC_Deprovisioning_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$script:DynamicGroupLogPath = ".\DynamicGroup_Users_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$script:DynamicGroupName = "W365-UserProvisioning-DYN"
$script:DynamicGroupUsers = @()

# Known Windows 365 licensing groups from the screenshot
$script:KnownLicensingGroups = @(
    "W365 - 2vCPU_8GB_128GBsku",
    "W365 - 4vCPU_16gbRAM_128gbSKU",
    "W365 - 4vCPU_16gbRAM_512gbSKU",
    "W365 - 2vCPU_4gb_128GBsku",
    "W365 - 2vCPU_8gbRAM_256gbSKU"
)

#region Helper Functions

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Console output with colors
    switch ($Level) {
        "INFO"    { Write-Host $logMessage -ForegroundColor Cyan }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        "ERROR"   { Write-Host $logMessage -ForegroundColor Red }
        "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
    }
    
    # File output
    Add-Content -Path $script:LogPath -Value $logMessage
}

function Show-Banner {
    Clear-Host
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "      Windows 365 Cloud PC User Deprovisioning Tool           " -ForegroundColor Yellow
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Connect-ToMicrosoftGraph {
    try {
        Write-Log "Connecting to Microsoft Graph..." -Level INFO
        
        # Required permissions for this script
        $requiredScopes = @(
            "CloudPC.ReadWrite.All",
            "Directory.Read.All",
            "Group.ReadWrite.All",
            "User.Read.All",
            "DeviceManagementConfiguration.ReadWrite.All"
        )
        
        Connect-MgGraph -Scopes $requiredScopes -NoWelcome
        
        $context = Get-MgContext
        Write-Log "Connected to Microsoft Graph successfully" -Level SUCCESS
        Write-Log "Account: $($context.Account)" -Level INFO
        Write-Log "Tenant: $($context.TenantId)" -Level INFO
        
        return $true
    }
    catch {
        Write-Log "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Get-UserGroups {
    param([string]$UserPrincipalName)
    
    try {
        Write-Log "Retrieving group memberships for user: $UserPrincipalName" -Level INFO
        
        $user = Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
        $groups = Get-MgUserMemberOf -UserId $user.Id -All | Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group' }
        
        $groupDetails = @()
        foreach ($group in $groups) {
            $groupInfo = Get-MgGroup -GroupId $group.Id
            $groupDetails += [PSCustomObject]@{
                Id = $groupInfo.Id
                DisplayName = $groupInfo.DisplayName
                GroupTypes = $groupInfo.GroupTypes
                IsDynamic = $groupInfo.GroupTypes -contains "DynamicMembership"
            }
        }
        
        Write-Log "Found $($groupDetails.Count) groups for user" -Level INFO
        return $groupDetails
    }
    catch {
        Write-Log "Error retrieving groups for user $UserPrincipalName : $($_.Exception.Message)" -Level ERROR
        return @()
    }
}

function Get-AllProvisioningPolicyGroups {
    try {
        Write-Log "Retrieving all Windows 365 provisioning policies..." -Level INFO
        
        # Get all provisioning policies
        $uri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies"
        $policies = Invoke-MgGraphRequest -Uri $uri -Method GET
        
        $allProvisioningGroups = @()
        
        if ($policies.value) {
            Write-Log "Found $($policies.value.Count) provisioning policy/policies" -Level INFO
            
            foreach ($policy in $policies.value) {
                Write-Log "Processing policy: $($policy.displayName)" -Level INFO
                
                # Get assignments for this policy
                $assignmentUri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies/$($policy.id)/assignments"
                $assignments = Invoke-MgGraphRequest -Uri $assignmentUri -Method GET
                
                if ($assignments.value) {
                    foreach ($assignment in $assignments.value) {
                        if ($assignment.target.'@odata.type' -eq '#microsoft.graph.groupAssignmentTarget') {
                            $groupId = $assignment.target.groupId
                            
                            try {
                                $group = Get-MgGroup -GroupId $groupId -ErrorAction Stop
                                $allProvisioningGroups += [PSCustomObject]@{
                                    GroupId = $group.Id
                                    GroupName = $group.DisplayName
                                    PolicyId = $policy.id
                                    PolicyName = $policy.displayName
                                    GroupTypes = $group.GroupTypes
                                    IsDynamic = $group.GroupTypes -contains "DynamicMembership"
                                }
                                Write-Log "  - Group: $($group.DisplayName)" -Level INFO
                            }
                            catch {
                                Write-Log "  - Warning: Could not retrieve group with ID: $groupId" -Level WARNING
                            }
                        }
                    }
                }
            }
        }
        else {
            Write-Log "No provisioning policies found" -Level WARNING
        }
        
        return $allProvisioningGroups
    }
    catch {
        Write-Log "Error retrieving provisioning policies: $($_.Exception.Message)" -Level ERROR
        return @()
    }
}

function Get-ProvisioningGroups {
    param(
        [array]$UserGroups,
        [array]$AllProvisioningPolicyGroups
    )
    
    $userProvisioningGroups = @()
    
    # Match user's groups against provisioning policy groups
    foreach ($userGroup in $UserGroups) {
        $matchingPolicyGroup = $AllProvisioningPolicyGroups | Where-Object { $_.GroupId -eq $userGroup.Id }
        
        if ($matchingPolicyGroup) {
            $userProvisioningGroups += [PSCustomObject]@{
                Id = $userGroup.Id
                DisplayName = $userGroup.DisplayName
                GroupTypes = $userGroup.GroupTypes
                IsDynamic = $userGroup.IsDynamic
                PolicyName = $matchingPolicyGroup.PolicyName
                PolicyId = $matchingPolicyGroup.PolicyId
            }
        }
    }
    
    return $userProvisioningGroups
}

function Get-LicensingGroups {
    param([array]$UserGroups)
    
    $licensingGroups = @()
    
    # Check against known licensing groups
    foreach ($group in $UserGroups) {
        if ($script:KnownLicensingGroups -contains $group.DisplayName) {
            $licensingGroups += $group
        }
    }
    
    # Also look for groups that match Windows 365 licensing patterns
    $licensingPatterns = @(
        "*W365*",
        "*Windows365*",
        "*CloudPC*",
        "*vCPU*",
        "*RAM*SKU*"
    )
    
    foreach ($group in $UserGroups) {
        if ($licensingGroups.DisplayName -notcontains $group.DisplayName) {
            foreach ($pattern in $licensingPatterns) {
                if ($group.DisplayName -like $pattern -and $group.DisplayName -notlike "*Provisioning*") {
                    $licensingGroups += $group
                    break
                }
            }
        }
    }
    
    return $licensingGroups | Sort-Object -Property DisplayName -Unique
}

function Remove-UserFromGroup {
    param(
        [string]$UserId,
        [string]$GroupId,
        [string]$GroupName,
        [bool]$IsDynamic
    )
    
    try {
        if ($IsDynamic) {
            Write-Log "Group '$GroupName' is a dynamic group - cannot remove user directly" -Level WARNING
            
            # Log this user for dynamic group report
            $script:DynamicGroupUsers += [PSCustomObject]@{
                UserPrincipalName = $UserId
                GroupName = $GroupName
                GroupId = $GroupId
                Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
            
            return $false
        }
        
        Write-Log "Removing user from group: $GroupName" -Level INFO
        Remove-MgGroupMemberByRef -GroupId $GroupId -DirectoryObjectId $UserId -ErrorAction Stop
        Write-Log "Successfully removed user from group: $GroupName" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Error removing user from group $GroupName : $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Get-UserCloudPCs {
    param([string]$UserPrincipalName)
    
    try {
        Write-Log "Retrieving Cloud PCs for user: $UserPrincipalName" -Level INFO
        
        # Get all Cloud PCs and filter by user
        $uri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/cloudPCs?`$filter=userPrincipalName eq '$UserPrincipalName'"
        $cloudPCs = Invoke-MgGraphRequest -Uri $uri -Method GET
        
        if ($cloudPCs.value) {
            Write-Log "Found $($cloudPCs.value.Count) Cloud PC(s) for user" -Level INFO
            return $cloudPCs.value
        }
        else {
            Write-Log "No Cloud PCs found for user" -Level WARNING
            return @()
        }
    }
    catch {
        Write-Log "Error retrieving Cloud PCs for user: $($_.Exception.Message)" -Level ERROR
        return @()
    }
}

function Test-CloudPCGracePeriod {
    param([object]$CloudPC)
    
    if ($CloudPC.status -eq "InGracePeriod" -or $CloudPC.gracePeriodEndDateTime) {
        return $true
    }
    return $false
}

function End-CloudPCGracePeriod {
    param([string]$CloudPCId)
    
    try {
        Write-Log "Ending grace period for Cloud PC: $CloudPCId" -Level INFO
        
        $uri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/cloudPCs/$CloudPCId/endGracePeriod"
        Invoke-MgGraphRequest -Uri $uri -Method POST
        
        Write-Log "Grace period ended successfully" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Error ending grace period: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Process-UserDeprovisioning {
    param([string]$UserPrincipalName)
    
    Write-Log "========================================" -Level INFO
    Write-Log "Processing user: $UserPrincipalName" -Level INFO
    Write-Log "========================================" -Level INFO
    
    try {
        # Verify user exists
        $user = Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
        Write-Log "User found: $($user.DisplayName)" -Level SUCCESS
        
        # Get user's groups
        $userGroups = Get-UserGroups -UserPrincipalName $UserPrincipalName
        
        if ($userGroups.Count -eq 0) {
            Write-Log "User is not a member of any groups" -Level WARNING
            return
        }
        
        # Identify provisioning groups
        $provisioningGroups = Get-ProvisioningGroups -UserGroups $userGroups
        Write-Log "Found $($provisioningGroups.Count) provisioning group(s)" -Level INFO
        
        # Identify licensing groups
        $licensingGroups = Get-LicensingGroups -UserGroups $userGroups
        Write-Log "Found $($licensingGroups.Count) licensing group(s)" -Level INFO
        
        # Check for dynamic group membership
        $dynamicGroupMember = $userGroups | Where-Object { $_.DisplayName -eq $script:DynamicGroupName }
        if ($dynamicGroupMember) {
            Write-Log "User is a member of dynamic group: $script:DynamicGroupName" -Level WARNING
            Write-Log "User will be logged to CSV file for manual review" -Level WARNING
            
            $script:DynamicGroupUsers += [PSCustomObject]@{
                UserPrincipalName = $UserPrincipalName
                GroupName = $script:DynamicGroupName
                GroupId = $dynamicGroupMember.Id
                Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
        
        # Remove user from provisioning groups
        Write-Log "Removing user from provisioning groups..." -Level INFO
        foreach ($group in $provisioningGroups) {
            Remove-UserFromGroup -UserId $user.Id -GroupId $group.Id -GroupName $group.DisplayName -IsDynamic $group.IsDynamic
        }
        
        # Remove user from licensing groups
        Write-Log "Removing user from licensing groups..." -Level INFO
        foreach ($group in $licensingGroups) {
            Remove-UserFromGroup -UserId $user.Id -GroupId $group.Id -GroupName $group.DisplayName -IsDynamic $group.IsDynamic
        }
        
        # Wait 3 minutes for changes to propagate
        Write-Log "Waiting 3 minutes for group membership changes to propagate..." -Level INFO
        Start-Sleep -Seconds 180
        
        # Check Cloud PC status
        $cloudPCs = Get-UserCloudPCs -UserPrincipalName $UserPrincipalName
        
        if ($cloudPCs.Count -eq 0) {
            Write-Log "No Cloud PCs found for user after group removal" -Level INFO
            return
        }
        
        # Check for grace period and offer to end it
        foreach ($cloudPC in $cloudPCs) {
            Write-Log "Cloud PC: $($cloudPC.displayName) - Status: $($cloudPC.status)" -Level INFO
            
            if (Test-CloudPCGracePeriod -CloudPC $cloudPC) {
                Write-Log "Cloud PC is in grace period" -Level WARNING
                Write-Log "Grace Period End Date: $($cloudPC.gracePeriodEndDateTime)" -Level INFO
                
                $response = Read-Host "Do you want to end the grace period for this Cloud PC? (Y/N)"
                if ($response -eq 'Y' -or $response -eq 'y') {
                    End-CloudPCGracePeriod -CloudPCId $cloudPC.id
                }
                else {
                    Write-Log "Grace period will remain active until: $($cloudPC.gracePeriodEndDateTime)" -Level INFO
                }
            }
            else {
                Write-Log "Cloud PC is not in grace period" -Level INFO
            }
        }
        
        Write-Log "Deprovisioning completed for user: $UserPrincipalName" -Level SUCCESS
    }
    catch {
        Write-Log "Error processing user $UserPrincipalName : $($_.Exception.Message)" -Level ERROR
    }
}

#endregion

#region Main Script

Show-Banner

Write-Log "Script started" -Level INFO
Write-Log "Log file: $script:LogPath" -Level INFO

# Connect to Microsoft Graph
if (-not (Connect-ToMicrosoftGraph)) {
    Write-Log "Failed to connect to Microsoft Graph. Exiting." -Level ERROR
    exit
}

# Determine input method if not specified
if (-not $InputMethod) {
    Write-Host ""
    Write-Host "Select input method:" -ForegroundColor Yellow
    Write-Host "1. Manual - Enter user UPN(s) manually" -ForegroundColor White
    Write-Host "2. CSV - Import users from CSV file" -ForegroundColor White
    Write-Host ""
    
    $choice = Read-Host "Enter choice (1 or 2)"
    
    switch ($choice) {
        "1" { $InputMethod = "Manual" }
        "2" { $InputMethod = "CSV" }
        default { 
            Write-Log "Invalid choice. Exiting." -Level ERROR
            exit
        }
    }
}

# Process users based on input method
switch ($InputMethod) {
    "Manual" {
        if (-not $UserPrincipalName) {
            Write-Host ""
            Write-Host "Enter user UPN (or multiple UPNs separated by commas):" -ForegroundColor Yellow
            $input = Read-Host "UPN(s)"
            $users = $input -split ',' | ForEach-Object { $_.Trim() }
        }
        else {
            $users = @($UserPrincipalName)
        }
        
        foreach ($upn in $users) {
            if ($upn) {
                Process-UserDeprovisioning -UserPrincipalName $upn
                Write-Host ""
            }
        }
    }
    
    "CSV" {
        if (-not $CsvPath) {
            Write-Host ""
            $CsvPath = Read-Host "Enter path to CSV file"
        }
        
        if (-not (Test-Path $CsvPath)) {
            Write-Log "CSV file not found: $CsvPath" -Level ERROR
            exit
        }
        
        try {
            $csvUsers = Import-Csv -Path $CsvPath
            
            if (-not $csvUsers[0].PSObject.Properties["UserPrincipalName"]) {
                Write-Log "CSV file must contain a 'UserPrincipalName' column" -Level ERROR
                exit
            }
            
            Write-Log "Found $($csvUsers.Count) user(s) in CSV file" -Level INFO
            
            foreach ($csvUser in $csvUsers) {
                Process-UserDeprovisioning -UserPrincipalName $csvUser.UserPrincipalName
                Write-Host ""
            }
        }
        catch {
            Write-Log "Error reading CSV file: $($_.Exception.Message)" -Level ERROR
            exit
        }
    }
}

# Export dynamic group users to CSV if any were found
if ($script:DynamicGroupUsers.Count -gt 0) {
    Write-Log "Exporting $($script:DynamicGroupUsers.Count) dynamic group user(s) to CSV" -Level INFO
    $script:DynamicGroupUsers | Export-Csv -Path $script:DynamicGroupLogPath -NoTypeInformation
    Write-Log "Dynamic group users saved to: $script:DynamicGroupLogPath" -Level SUCCESS
}

Write-Log "Script completed" -Level SUCCESS
Write-Log "Log file: $script:LogPath" -Level INFO

# Disconnect from Microsoft Graph
Disconnect-MgGraph | Out-Null
Write-Log "Disconnected from Microsoft Graph" -Level INFO

#endregion
