# Cloud PC Recovery & Reprovisioning Tool
# Requires Microsoft.Graph PowerShell Module

param(
    [Parameter(Mandatory=$false)]
    [string]$TenantId,
    
    [Parameter(Mandatory=$false)]
    [string]$ClientId,
    
    [Parameter(Mandatory=$false)]
    [string]$ClientSecret
)

# Import required modules
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module Microsoft.Graph.Groups -ErrorAction Stop
Import-Module Microsoft.Graph.DeviceManagement.Administration -ErrorAction Stop

# Global variables
$Global:AuthenticationMethod = $null
$Global:ConnectedToGraph = $false

# Function to display banner
function Show-Banner {
    Clear-Host
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host "   Cloud PC Recovery & Reprovisioning Tool    " -ForegroundColor Yellow
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host ""
}

# Function to authenticate to Microsoft Graph
function Connect-ToMicrosoftGraph {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret
    )
    
    try {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
        
        if ($TenantId -and $ClientId -and $ClientSecret) {
            # App-only authentication
            $secureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential($ClientId, $secureSecret)
            Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $credential -NoWelcome
            $Global:AuthenticationMethod = "App-only"
        } else {
            # Interactive authentication
            Connect-MgGraph -Scopes "DeviceManagementConfiguration.ReadWrite.All", "Group.ReadWrite.All", "Directory.Read.All" -NoWelcome
            $Global:AuthenticationMethod = "Interactive"
        }
        
        $Global:ConnectedToGraph = $true
        Write-Host "Successfully connected to Microsoft Graph using $($Global:AuthenticationMethod) authentication" -ForegroundColor Green
        Write-Host ""
        return $true
    }
    catch {
        Write-Host "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to get all Windows 365 provisioning policies
function Get-CloudPCProvisioningPolicies {
    try {
        Write-Host "Retrieving Windows 365 Provisioning Policies..." -ForegroundColor Yellow
        
        $uri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies"
        $policies = Invoke-MgGraphRequest -Uri $uri -Method GET
        
        if ($policies.value.Count -eq 0) {
            Write-Host "No provisioning policies found" -ForegroundColor Yellow
            return $null
        }
        
        Write-Host "Found $($policies.value.Count) provisioning policies:" -ForegroundColor Green
        
        $policyList = @()
        foreach ($policy in $policies.value) {
            $policyInfo = [PSCustomObject]@{
                Id = $policy.id
                DisplayName = $policy.displayName
                Description = $policy.description
                CloudPCSize = $policy.cloudPcSize
                ImageType = $policy.imageType
                AssignedGroups = @()
            }
            
            # Get group assignments for each policy
            $assignmentUri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies/$($policy.id)/assignments"
            $assignments = Invoke-MgGraphRequest -Uri $assignmentUri -Method GET
            
            foreach ($assignment in $assignments.value) {
                if ($assignment.target.'@odata.type' -eq '#microsoft.graph.groupAssignmentTarget') {
                    try {
                        $group = Get-MgGroup -GroupId $assignment.target.groupId -ErrorAction SilentlyContinue
                        if ($group) {
                            $policyInfo.AssignedGroups += [PSCustomObject]@{
                                GroupId = $group.Id
                                GroupName = $group.DisplayName
                                GroupType = if ($group.GroupTypes -contains "DynamicMembership") { "Dynamic" } else { "Static" }
                            }
                        }
                    }
                    catch {
                        $policyInfo.AssignedGroups += [PSCustomObject]@{
                            GroupId = $assignment.target.groupId
                            GroupName = "Unknown (Access Denied or Deleted)"
                            GroupType = "Unknown"
                        }
                    }
                }
            }
            
            $policyList += $policyInfo
            Write-Host "  - $($policy.displayName) (Groups: $($policyInfo.AssignedGroups.Count))" -ForegroundColor White
        }
        
        return $policyList
    }
    catch {
        Write-Host "Error retrieving provisioning policies: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

# Function to display policies in a formatted table
function Show-ProvisioningPolicies {
    param([array]$Policies)
    
    if (-not $Policies) {
        Write-Host "No policies to display" -ForegroundColor Yellow
        return
    }
    
    Write-Host "`nProvisioning Policies and Group Assignments:" -ForegroundColor Cyan
    Write-Host "=============================================" -ForegroundColor Cyan
    
    for ($i = 0; $i -lt $Policies.Count; $i++) {
        $policy = $Policies[$i]
        Write-Host "`n[$($i + 1)] $($policy.DisplayName)" -ForegroundColor Yellow
        Write-Host "    Policy ID: $($policy.Id)" -ForegroundColor Gray
        Write-Host "    Description: $($policy.Description)" -ForegroundColor Gray
        Write-Host "    Cloud PC Size: $($policy.CloudPCSize)" -ForegroundColor Gray
        Write-Host "    Assigned Groups:" -ForegroundColor White
        
        if ($policy.AssignedGroups.Count -eq 0) {
            Write-Host "      No groups assigned" -ForegroundColor Yellow
        } else {
            foreach ($group in $policy.AssignedGroups) {
                Write-Host "      - $($group.GroupName) ($($group.GroupType))" -ForegroundColor White
                Write-Host "        Group ID: $($group.GroupId)" -ForegroundColor Gray
            }
        }
    }
}

# Function to move user between groups
function Move-UserBetweenGroups {
    param(
        [string]$UserPrincipalName,
        [string]$SourceGroupId,
        [string]$TargetGroupId
    )
    
    try {
        # Get user object
        Write-Host "Looking up user: $UserPrincipalName..." -ForegroundColor Yellow
        $user = Get-MgUser -Filter "userPrincipalName eq '$UserPrincipalName'" -ErrorAction Stop
        
        if (-not $user) {
            Write-Host "User not found: $UserPrincipalName" -ForegroundColor Red
            return $false
        }
        
        Write-Host "Found user: $($user.DisplayName) ($($user.Id))" -ForegroundColor Green
        
        # Remove from source group
        if ($SourceGroupId) {
            Write-Host "Removing user from source group..." -ForegroundColor Yellow
            try {
                Remove-MgGroupMember -GroupId $SourceGroupId -DirectoryObjectId $user.Id -ErrorAction Stop
                Write-Host "Successfully removed user from source group" -ForegroundColor Green
            }
            catch {
                if ($_.Exception.Message -like "*does not exist*") {
                    Write-Host "User was not a member of the source group" -ForegroundColor Yellow
                } else {
                    Write-Host "Warning: Could not remove user from source group: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
        }
        
        # Add to target group
        Write-Host "Adding user to target group..." -ForegroundColor Yellow
        try {
            New-MgGroupMember -GroupId $TargetGroupId -DirectoryObjectId $user.Id -ErrorAction Stop
            Write-Host "Successfully added user to target group" -ForegroundColor Green
            return $true
        }
        catch {
            if ($_.Exception.Message -like "*already exists*") {
                Write-Host "User is already a member of the target group" -ForegroundColor Yellow
                return $true
            } else {
                Write-Host "Error adding user to target group: $($_.Exception.Message)" -ForegroundColor Red
                return $false
            }
        }
    }
    catch {
        Write-Host "Error processing user move: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to get user's Cloud PC status
function Get-UserCloudPCStatus {
    param([string]$UserPrincipalName)
    
    try {
        Write-Host "Checking Cloud PC status for: $UserPrincipalName..." -ForegroundColor Yellow
        
        # Get user object
        $user = Get-MgUser -Filter "userPrincipalName eq '$UserPrincipalName'" -ErrorAction Stop
        
        # Get Cloud PCs for the user
        $uri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/cloudPCs?`$filter=userPrincipalName eq '$UserPrincipalName'"
        $cloudPCs = Invoke-MgGraphRequest -Uri $uri -Method GET
        
        if ($cloudPCs.value.Count -eq 0) {
            Write-Host "No Cloud PCs found for user: $UserPrincipalName" -ForegroundColor Yellow
            return $null
        }
        
        $cloudPCList = @()
        foreach ($cloudPC in $cloudPCs.value) {
            $cloudPCInfo = [PSCustomObject]@{
                Id = $cloudPC.id
                DisplayName = $cloudPC.displayName
                Status = $cloudPC.status
                ProvisioningPolicyId = $cloudPC.provisioningPolicyId
                ProvisioningPolicyName = "Unknown"
                GracePeriodEndDateTime = $cloudPC.gracePeriodEndDateTime
                IsInGracePeriod = $false
                LastModifiedDateTime = $cloudPC.lastModifiedDateTime
                UserPrincipalName = $cloudPC.userPrincipalName
            }
            
            # Check if in grace period
            if ($cloudPC.gracePeriodEndDateTime) {
                $gracePeriodEnd = [DateTime]::Parse($cloudPC.gracePeriodEndDateTime)
                $cloudPCInfo.IsInGracePeriod = $gracePeriodEnd -gt (Get-Date)
            }
            
            # Get provisioning policy name
            if ($cloudPC.provisioningPolicyId) {
                try {
                    $policyUri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies/$($cloudPC.provisioningPolicyId)"
                    $policy = Invoke-MgGraphRequest -Uri $policyUri -Method GET
                    $cloudPCInfo.ProvisioningPolicyName = $policy.displayName
                }
                catch {
                    $cloudPCInfo.ProvisioningPolicyName = "Policy not found or access denied"
                }
            }
            
            $cloudPCList += $cloudPCInfo
        }
        
        return $cloudPCList
    }
    catch {
        Write-Host "Error getting Cloud PC status: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

# Function to display Cloud PC status
function Show-CloudPCStatus {
    param([array]$CloudPCs)
    
    if (-not $CloudPCs -or $CloudPCs.Count -eq 0) {
        Write-Host "No Cloud PCs to display" -ForegroundColor Yellow
        return
    }
    
    Write-Host "`nCloud PC Status:" -ForegroundColor Cyan
    Write-Host "================" -ForegroundColor Cyan
    
    foreach ($cloudPC in $CloudPCs) {
        Write-Host "`nCloud PC: $($cloudPC.DisplayName)" -ForegroundColor Yellow
        Write-Host "  Status: $($cloudPC.Status)" -ForegroundColor White
        Write-Host "  Provisioning Policy: $($cloudPC.ProvisioningPolicyName)" -ForegroundColor White
        Write-Host "  User: $($cloudPC.UserPrincipalName)" -ForegroundColor White
        Write-Host "  Last Modified: $($cloudPC.LastModifiedDateTime)" -ForegroundColor Gray
        
        if ($cloudPC.IsInGracePeriod) {
            Write-Host "  Grace Period: Active (Ends: $($cloudPC.GracePeriodEndDateTime))" -ForegroundColor Red
        } elseif ($cloudPC.GracePeriodEndDateTime) {
            Write-Host "  Grace Period: Ended ($($cloudPC.GracePeriodEndDateTime))" -ForegroundColor Green
        } else {
            Write-Host "  Grace Period: Not applicable" -ForegroundColor Gray
        }
    }
}

# Function to end grace period
function End-CloudPCGracePeriod {
    param([string]$CloudPCId)
    
    try {
        Write-Host "Ending grace period for Cloud PC: $CloudPCId..." -ForegroundColor Yellow
        
        $uri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/cloudPCs/$CloudPCId/endGracePeriod"
        Invoke-MgGraphRequest -Uri $uri -Method POST
        
        Write-Host "Grace period ended successfully" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Error ending grace period: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to wait for provisioning
function Wait-ForProvisioning {
    param(
        [string]$UserPrincipalName,
        [int]$TimeoutMinutes = 30
    )
    
    Write-Host "Monitoring provisioning status for $UserPrincipalName..." -ForegroundColor Yellow
    Write-Host "Timeout set to $TimeoutMinutes minutes" -ForegroundColor Gray
    
    $startTime = Get-Date
    $timeout = $startTime.AddMinutes($TimeoutMinutes)
    
    do {
        $cloudPCs = Get-UserCloudPCStatus -UserPrincipalName $UserPrincipalName
        
        if ($cloudPCs) {
            $provisionedPC = $cloudPCs | Where-Object { $_.Status -eq "Provisioned" }
            if ($provisionedPC) {
                Write-Host "`nProvisioning completed!" -ForegroundColor Green
                Show-CloudPCStatus -CloudPCs @($provisionedPC)
                return $provisionedPC
            }
        }
        
        $elapsed = (Get-Date) - $startTime
        Write-Host "  Elapsed: $([math]::Round($elapsed.TotalMinutes, 1)) minutes - Status: Provisioning..." -ForegroundColor Yellow
        Start-Sleep -Seconds 30
        
    } while ((Get-Date) -lt $timeout)
    
    Write-Host "`nTimeout reached. Provisioning may still be in progress." -ForegroundColor Yellow
    return $null
}

# Main menu function
function Show-MainMenu {
    Write-Host "`nMain Menu:" -ForegroundColor Cyan
    Write-Host "==========" -ForegroundColor Cyan
    Write-Host "1. View all Provisioning Policies and Group Assignments"
    Write-Host "2. Move User Between Provisioning Policies"
    Write-Host "3. Check User Cloud PC Status"
    Write-Host "4. End Grace Period for Cloud PC"
    Write-Host "5. Monitor New Cloud PC Provisioning"
    Write-Host "6. Complete Recovery Workflow"
    Write-Host "7. Disconnect and Exit"
    Write-Host ""
}

# Complete recovery workflow
function Start-RecoveryWorkflow {
    Write-Host "`nStarting Complete Recovery Workflow" -ForegroundColor Cyan
    Write-Host "====================================" -ForegroundColor Cyan
    
    # Step 1: Get user
    $userPrincipalName = Read-Host "Enter user's email address (UPN)"
    if (-not $userPrincipalName) {
        Write-Host "Invalid user principal name" -ForegroundColor Red
        return
    }
    
    # Step 2: Get policies
    $policies = Get-CloudPCProvisioningPolicies
    if (-not $policies) {
        Write-Host "Cannot proceed without provisioning policies" -ForegroundColor Red
        return
    }
    
    Show-ProvisioningPolicies -Policies $policies
    
    # Step 3: Select source policy
    Write-Host "`nSelect source provisioning policy (or press Enter to skip removal):" -ForegroundColor Yellow
    $sourceChoice = Read-Host "Enter policy number (1-$($policies.Count))"
    
    $sourceGroupId = $null
    if ($sourceChoice -and [int]$sourceChoice -ge 1 -and [int]$sourceChoice -le $policies.Count) {
        $sourcePolicy = $policies[[int]$sourceChoice - 1]
        if ($sourcePolicy.AssignedGroups.Count -eq 1) {
            $sourceGroupId = $sourcePolicy.AssignedGroups[0].GroupId
        } elseif ($sourcePolicy.AssignedGroups.Count -gt 1) {
            Write-Host "Multiple groups found for source policy. Select group:" -ForegroundColor Yellow
            for ($i = 0; $i -lt $sourcePolicy.AssignedGroups.Count; $i++) {
                Write-Host "  [$($i + 1)] $($sourcePolicy.AssignedGroups[$i].GroupName)"
            }
            $groupChoice = Read-Host "Enter group number"
            if ($groupChoice -and [int]$groupChoice -ge 1 -and [int]$groupChoice -le $sourcePolicy.AssignedGroups.Count) {
                $sourceGroupId = $sourcePolicy.AssignedGroups[[int]$groupChoice - 1].GroupId
            }
        }
    }
    
    # Step 4: Select target policy
    Write-Host "`nSelect target (recovery) provisioning policy:" -ForegroundColor Yellow
    $targetChoice = Read-Host "Enter policy number (1-$($policies.Count))"
    
    if (-not $targetChoice -or [int]$targetChoice -lt 1 -or [int]$targetChoice -gt $policies.Count) {
        Write-Host "Invalid target policy selection" -ForegroundColor Red
        return
    }
    
    $targetPolicy = $policies[[int]$targetChoice - 1]
    $targetGroupId = $null
    
    if ($targetPolicy.AssignedGroups.Count -eq 1) {
        $targetGroupId = $targetPolicy.AssignedGroups[0].GroupId
    } elseif ($targetPolicy.AssignedGroups.Count -gt 1) {
        Write-Host "Multiple groups found for target policy. Select group:" -ForegroundColor Yellow
        for ($i = 0; $i -lt $targetPolicy.AssignedGroups.Count; $i++) {
            Write-Host "  [$($i + 1)] $($targetPolicy.AssignedGroups[$i].GroupName)"
        }
        $groupChoice = Read-Host "Enter group number"
        if ($groupChoice -and [int]$groupChoice -ge 1 -and [int]$groupChoice -le $targetPolicy.AssignedGroups.Count) {
            $targetGroupId = $targetPolicy.AssignedGroups[[int]$groupChoice - 1].GroupId
        }
    } else {
        Write-Host "No groups assigned to target policy" -ForegroundColor Red
        return
    }
    
    # Step 5: Move user
    Write-Host "`nMoving user between groups..." -ForegroundColor Yellow
    $moveResult = Move-UserBetweenGroups -UserPrincipalName $userPrincipalName -SourceGroupId $sourceGroupId -TargetGroupId $targetGroupId
    
    if (-not $moveResult) {
        Write-Host "Failed to move user. Workflow stopped." -ForegroundColor Red
        return
    }
    
    # Step 6: Check current Cloud PC status
    Write-Host "`nChecking current Cloud PC status..." -ForegroundColor Yellow
    $cloudPCs = Get-UserCloudPCStatus -UserPrincipalName $userPrincipalName
    
    if ($cloudPCs) {
        Show-CloudPCStatus -CloudPCs $cloudPCs
        
        # Step 7: Handle grace period
        $gracePeriodPCs = $cloudPCs | Where-Object { $_.IsInGracePeriod }
        if ($gracePeriodPCs) {
            Write-Host "`nFound Cloud PCs in grace period." -ForegroundColor Yellow
            $endGrace = Read-Host "Do you want to end the grace period to start reprovisioning? (y/n)"
            
            if ($endGrace -eq 'y' -or $endGrace -eq 'Y') {
                foreach ($pc in $gracePeriodPCs) {
                    End-CloudPCGracePeriod -CloudPCId $pc.Id
                }
            }
        }
    }
    
    # Step 8: Monitor new provisioning
    Write-Host "`nWould you like to monitor the new Cloud PC provisioning?" -ForegroundColor Yellow
    $monitor = Read-Host "Monitor provisioning? (y/n)"
    
    if ($monitor -eq 'y' -or $monitor -eq 'Y') {
        $timeoutMinutes = Read-Host "Enter timeout in minutes (default: 30)"
        if (-not $timeoutMinutes -or -not [int]::TryParse($timeoutMinutes, [ref]$null)) {
            $timeoutMinutes = 30
        }
        
        Wait-ForProvisioning -UserPrincipalName $userPrincipalName -TimeoutMinutes ([int]$timeoutMinutes)
    }
    
    Write-Host "`nRecovery workflow completed!" -ForegroundColor Green
}

# Main script execution
Show-Banner

# Check if required modules are available
try {
    Get-Module Microsoft.Graph.Authentication -ListAvailable | Out-Null
}
catch {
    Write-Host "Microsoft Graph PowerShell modules are required but not found." -ForegroundColor Red
    Write-Host "Please install using: Install-Module Microsoft.Graph -Scope CurrentUser" -ForegroundColor Yellow
    exit 1
}

# Connect to Microsoft Graph
if (-not (Connect-ToMicrosoftGraph -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret)) {
    exit 1
}

# Main application loop
do {
    Show-MainMenu
    $choice = Read-Host "Select an option (1-7)"
    
    switch ($choice) {
        "1" {
            $policies = Get-CloudPCProvisioningPolicies
            Show-ProvisioningPolicies -Policies $policies
            Read-Host "`nPress Enter to continue"
        }
        
        "2" {
            Write-Host "`nMove User Between Provisioning Policies" -ForegroundColor Cyan
            Write-Host "=======================================" -ForegroundColor Cyan
            
            $policies = Get-CloudPCProvisioningPolicies
            if ($policies) {
                Show-ProvisioningPolicies -Policies $policies
                
                $userPrincipalName = Read-Host "`nEnter user's email address (UPN)"
                
                Write-Host "`nSelect source group (or press Enter to skip):"
                $sourceChoice = Read-Host "Enter policy number (1-$($policies.Count))"
                
                Write-Host "`nSelect target group:"
                $targetChoice = Read-Host "Enter policy number (1-$($policies.Count))"
                
                if ($targetChoice -and [int]$targetChoice -ge 1 -and [int]$targetChoice -le $policies.Count) {
                    $sourceGroupId = $null
                    if ($sourceChoice -and [int]$sourceChoice -ge 1 -and [int]$sourceChoice -le $policies.Count) {
                        $sourcePolicy = $policies[[int]$sourceChoice - 1]
                        if ($sourcePolicy.AssignedGroups.Count -gt 0) {
                            $sourceGroupId = $sourcePolicy.AssignedGroups[0].GroupId
                        }
                    }
                    
                    $targetPolicy = $policies[[int]$targetChoice - 1]
                    if ($targetPolicy.AssignedGroups.Count -gt 0) {
                        $targetGroupId = $targetPolicy.AssignedGroups[0].GroupId
                        Move-UserBetweenGroups -UserPrincipalName $userPrincipalName -SourceGroupId $sourceGroupId -TargetGroupId $targetGroupId
                    } else {
                        Write-Host "Target policy has no assigned groups" -ForegroundColor Red
                    }
                } else {
                    Write-Host "Invalid selection" -ForegroundColor Red
                }
            }
            Read-Host "`nPress Enter to continue"
        }
        
        "3" {
            Write-Host "`nCheck User Cloud PC Status" -ForegroundColor Cyan
            Write-Host "==========================" -ForegroundColor Cyan
            
            $userPrincipalName = Read-Host "Enter user's email address (UPN)"
            $cloudPCs = Get-UserCloudPCStatus -UserPrincipalName $userPrincipalName
            Show-CloudPCStatus -CloudPCs $cloudPCs
            Read-Host "`nPress Enter to continue"
        }
        
        "4" {
            Write-Host "`nEnd Grace Period for Cloud PC" -ForegroundColor Cyan
            Write-Host "=============================" -ForegroundColor Cyan
            
            $userPrincipalName = Read-Host "Enter user's email address (UPN)"
            $cloudPCs = Get-UserCloudPCStatus -UserPrincipalName $userPrincipalName
            
            if ($cloudPCs) {
                Show-CloudPCStatus -CloudPCs $cloudPCs
                $gracePeriodPCs = $cloudPCs | Where-Object { $_.IsInGracePeriod }
                
                if ($gracePeriodPCs) {
                    foreach ($pc in $gracePeriodPCs) {
                        $confirm = Read-Host "`nEnd grace period for '$($pc.DisplayName)'? (y/n)"
                        if ($confirm -eq 'y' -or $confirm -eq 'Y') {
                            End-CloudPCGracePeriod -CloudPCId $pc.Id
                        }
                    }
                } else {
                    Write-Host "No Cloud PCs in grace period found" -ForegroundColor Yellow
                }
            }
            Read-Host "`nPress Enter to continue"
        }
        
        "5" {
            Write-Host "`nMonitor New Cloud PC Provisioning" -ForegroundColor Cyan
            Write-Host "==================================" -ForegroundColor Cyan
            
            $userPrincipalName = Read-Host "Enter user's email address (UPN)"
            $timeoutMinutes = Read-Host "Enter timeout in minutes (default: 30)"
            if (-not $timeoutMinutes -or -not [int]::TryParse($timeoutMinutes, [ref]$null)) {
                $timeoutMinutes = 30
            }
            
            Wait-ForProvisioning -UserPrincipalName $userPrincipalName -TimeoutMinutes ([int]$timeoutMinutes)
            Read-Host "`nPress Enter to continue"
        }
        
        "6" {
            Start-RecoveryWorkflow
            Read-Host "`nPress Enter to continue"
        }
        
        "7" {
            Write-Host "`nDisconnecting from Microsoft Graph..." -ForegroundColor Yellow
            Disconnect-MgGraph
            Write-Host "Goodbye!" -ForegroundColor Green
            exit
        }
        
        default {
            Write-Host "Invalid option. Please select 1-7." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
    
} while ($choice -ne "7")
