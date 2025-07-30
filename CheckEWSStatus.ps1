# Exchange Mailbox EWS and Calendar Permissions Checker
# This script reads mailboxes from a CSV file, checks EWS status, and verifies calendar permissions

param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath,
    
    [Parameter(Mandatory=$false)]
    [string]$OTDAccount = "otd@domain.com",
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "MailboxReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Function to check if OTD account has calendar permissions on a mailbox
function Get-OTDCalendarPermissions {
    param(
        [string]$MailboxIdentity,
        [string]$OTDAccount
    )
    
    try {
        # Try different calendar folder names for different languages
        $calendarPaths = @(
            "${MailboxIdentity}:\Calendar",
            "${MailboxIdentity}:\Kalender",  # German
            "${MailboxIdentity}:\Calendrier", # French
            "${MailboxIdentity}:\Calendario"  # Spanish/Italian
        )
        
        $permissions = $null
        $workingPath = $null
        
        # First, try to find the correct calendar path and check for OTD permissions
        foreach ($path in $calendarPaths) {
            try {
                # Check if this calendar path exists by getting all permissions
                $allPermissions = Get-MailboxFolderPermission -Identity $path -ErrorAction Stop
                $workingPath = $path
                
                # Look for the OTD account specifically
                $permissions = $allPermissions | Where-Object { $_.User.ToString() -eq $OTDAccount -or $_.User.ADRecipient.PrimarySmtpAddress -eq $OTDAccount }
                
                if ($permissions) {
                    break
                }
            }
            catch {
                continue
            }
        }
        
        if ($permissions) {
            return @{
                HasPermission = $true
                AccessRights = ($permissions.AccessRights -join ", ")
                FolderPath = $permissions.FolderName
                CalendarPath = $workingPath
            }
        }
        else {
            # No permissions found, but we have a working calendar path
            return @{
                HasPermission = $false
                AccessRights = "No permissions found for $OTDAccount"
                FolderPath = "N/A"
                CalendarPath = $workingPath
            }
        }
    }
    catch {
        return @{
            HasPermission = $false
            AccessRights = "Error: $($_.Exception.Message)"
            FolderPath = "N/A"
            CalendarPath = $null
        }
    }
}

# Function to prompt for permission assignment
function Prompt-ForOTDPermissionAssignment {
    param(
        [string]$MailboxIdentity,
        [string]$OTDAccount
    )
    
    do {
        $response = Read-Host "Grant $OTDAccount access to $MailboxIdentity's calendar? (y/n)"
        $response = $response.ToLower().Trim()
    } while ($response -ne 'y' -and $response -ne 'n' -and $response -ne 'yes' -and $response -ne 'no')
    
    return ($response -eq 'y' -or $response -eq 'yes')
}

# Function to select permission level
function Select-PermissionLevel {
    Write-Host ""
    Write-Host "Available Calendar Permission Levels:" -ForegroundColor Yellow
    Write-Host "1. Reviewer (Read-only access to calendar items)"
    Write-Host "2. Author (Read, create, and modify own calendar items)"
    Write-Host "3. Editor (Read, create, modify, and delete all calendar items)"
    Write-Host "4. Owner (Full control over the calendar)"
    Write-Host "5. PublishingEditor (Editor + create subfolders)"
    Write-Host "6. PublishingAuthor (Author + create subfolders)"
    Write-Host "7. NonEditingAuthor (Read and create items, modify/delete own items)"
    Write-Host "8. Contributor (Create items only)"
    Write-Host "9. AvailabilityOnly (View free/busy information only)"
    Write-Host "10. LimitedDetails (View free/busy and subject/location)"
    Write-Host ""
    
    do {
        $choice = Read-Host "Select permission level (1-10)"
        $validChoice = $choice -match '^([1-9]|10)

# Main script execution
try {
    Write-Host "Starting Exchange Mailbox Analysis..." -ForegroundColor Green
    Write-Host "CSV File: $CsvPath" -ForegroundColor Yellow
    Write-Host "Checking if OTD account has access: $OTDAccount" -ForegroundColor Yellow
    Write-Host "Output will be saved to: $OutputPath" -ForegroundColor Yellow
    Write-Host ""

    # Check if CSV file exists
    if (-not (Test-Path $CsvPath)) {
        throw "CSV file not found: $CsvPath"
    }

    # Read the CSV file
    # Assuming the CSV has a column named "Mailbox", "EmailAddress", "Identity", or "UserPrincipalName"
    $mailboxList = Import-Csv $CsvPath
    
    # Auto-detect the column name for mailbox identities
    $mailboxColumn = $null
    $possibleColumns = @("Mailbox", "EmailAddress", "Identity", "UserPrincipalName", "Email", "PrimarySmtpAddress")
    
    foreach ($col in $possibleColumns) {
        if ($mailboxList[0].PSObject.Properties.Name -contains $col) {
            $mailboxColumn = $col
            break
        }
    }
    
    if (-not $mailboxColumn) {
        Write-Host "Available columns in CSV:" -ForegroundColor Red
        $mailboxList[0].PSObject.Properties.Name | ForEach-Object { Write-Host "  - $_" }
        throw "Could not find a mailbox column. Please ensure your CSV has one of these columns: $($possibleColumns -join ', ')"
    }
    
    Write-Host "Using column '$mailboxColumn' for mailbox identities" -ForegroundColor Green
    Write-Host ""

    # Initialize results array
    $results = @()
    $totalMailboxes = $mailboxList.Count
    $currentCount = 0

    # Process each mailbox
    foreach ($row in $mailboxList) {
        $currentCount++
        $mailboxIdentity = $row.$mailboxColumn
        
        Write-Progress -Activity "Processing Mailboxes" -Status "Checking $mailboxIdentity ($currentCount of $totalMailboxes)" -PercentComplete (($currentCount / $totalMailboxes) * 100)
        
        Write-Host "Processing: $mailboxIdentity" -ForegroundColor Cyan
        
        try {
            # Check EWS status
            $casMailbox = Get-CASMailbox -Identity $mailboxIdentity -ErrorAction Stop
            $ewsEnabled = $casMailbox.EWSEnabled
            
            # Check if OTD account has calendar permissions on this mailbox
            $otdCalendarPerms = Get-OTDCalendarPermissions -MailboxIdentity $mailboxIdentity -OTDAccount $OTDAccount
            
            Write-Host "  EWS Enabled: $ewsEnabled" -ForegroundColor $(if ($ewsEnabled) { "Green" } else { "Red" })
            Write-Host "  OTD Calendar Access: $($otdCalendarPerms.HasPermission)" -ForegroundColor $(if ($otdCalendarPerms.HasPermission) { "Green" } else { "Yellow" })
            
            $permissionGranted = $false
            $grantedPermissionLevel = ""
            $permissionGrantMessage = ""
            
            if ($otdCalendarPerms.HasPermission) {
                Write-Host "  OTD Access Rights: $($otdCalendarPerms.AccessRights)" -ForegroundColor Green
            }
            else {
                # Prompt user if they want to grant OTD access to this mailbox's calendar
                if ($otdCalendarPerms.CalendarPath) {
                    Write-Host "  $OTDAccount does not have access to $mailboxIdentity's calendar" -ForegroundColor Yellow
                    
                    if (Prompt-ForOTDPermissionAssignment -MailboxIdentity $mailboxIdentity -OTDAccount $OTDAccount) {
                        $permissionLevel = Select-PermissionLevel
                        Write-Host "  Granting '$permissionLevel' permission to $OTDAccount for $mailboxIdentity's calendar..." -ForegroundColor Cyan
                        
                        $grantResult = Grant-OTDCalendarPermission -CalendarPath $otdCalendarPerms.CalendarPath -OTDAccount $OTDAccount -PermissionLevel $permissionLevel
                        
                        if ($grantResult.Success) {
                            Write-Host "  $($grantResult.Message)" -ForegroundColor Green
                            $permissionGranted = $true
                            $grantedPermissionLevel = $permissionLevel
                            $permissionGrantMessage = $grantResult.Message
                            
                            # Update calendar permissions status
                            $otdCalendarPerms.HasPermission = $true
                            $otdCalendarPerms.AccessRights = $permissionLevel
                        }
                        else {
                            Write-Host "  $($grantResult.Message)" -ForegroundColor Red
                            $permissionGrantMessage = $grantResult.Message
                        }
                    }
                    else {
                        Write-Host "  Permission not granted - skipping..." -ForegroundColor Gray
                        $permissionGrantMessage = "User declined to grant permission"
                    }
                }
                else {
                    Write-Host "  Cannot find calendar path - unable to grant permissions" -ForegroundColor Red
                    $permissionGrantMessage = "Calendar path not found"
                }
            }
            
            # Create result object
            $result = [PSCustomObject]@{
                Mailbox = $mailboxIdentity
                EWSEnabled = $ewsEnabled
                OTDHasCalendarAccess = $otdCalendarPerms.HasPermission
                OTDCalendarAccessRights = $otdCalendarPerms.AccessRights
                CalendarFolderPath = $otdCalendarPerms.FolderPath
                PermissionGrantedToOTD = $permissionGranted
                GrantedPermissionLevel = $grantedPermissionLevel
                PermissionGrantMessage = $permissionGrantMessage
                Status = "Success"
                Error = ""
            }
        }
        catch {
            Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
            
            $result = [PSCustomObject]@{
                Mailbox = $mailboxIdentity
                EWSEnabled = "Error"
                OTDHasCalendarAccess = $false
                OTDCalendarAccessRights = "Error retrieving permissions"
                CalendarFolderPath = "N/A"
                PermissionGrantedToOTD = $false
                GrantedPermissionLevel = ""
                PermissionGrantMessage = "Error occurred during processing"
                Status = "Failed"
                Error = $_.Exception.Message
            }
        }
        
        $results += $result
        Write-Host ""
    }

    # Export results to CSV
    $results | Export-Csv -Path $OutputPath -NoTypeInformation
    
    # Display summary
    Write-Host "Analysis Complete!" -ForegroundColor Green
    Write-Host "Results saved to: $OutputPath" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Summary:" -ForegroundColor Cyan
    Write-Host "Total Mailboxes Processed: $($results.Count)"
    Write-Host "EWS Enabled: $(($results | Where-Object { $_.EWSEnabled -eq $true }).Count)"
    Write-Host "EWS Disabled: $(($results | Where-Object { $_.EWSEnabled -eq $false }).Count)"
    Write-Host "OTD Has Calendar Access: $(($results | Where-Object { $_.OTDHasCalendarAccess -eq $true }).Count)"
    Write-Host "Permissions Granted to OTD: $(($results | Where-Object { $_.PermissionGrantedToOTD -eq $true }).Count)"
    Write-Host "Errors: $(($results | Where-Object { $_.Status -eq 'Failed' }).Count)"
    
}
catch {
    Write-Error "Script failed: $($_.Exception.Message)"
    exit 1
}

# Example CSV format (save as mailboxes.csv):
<#
Mailbox
user1@domain.com
user2@domain.com
user3@domain.com
#>

# Usage examples:
# .\CheckMailboxEWSAndPermissions.ps1 -CsvPath "C:\temp\mailboxes.csv"
# .\CheckMailboxEWSAndPermissions.ps1 -CsvPath "C:\temp\mailboxes.csv" -OTDAccount "admin@domain.com"
# .\CheckMailboxEWSAndPermissions.ps1 -CsvPath "C:\temp\mailboxes.csv" -OTDAccount "otd@domain.com" -OutputPath "C:\reports\mailbox_report.csv"
        if (-not $validChoice) {
            Write-Host "Invalid choice. Please enter a number between 1 and 10." -ForegroundColor Red
        }
    } while (-not $validChoice)
    
    $permissionLevels = @{
        1 = "Reviewer"
        2 = "Author"
        3 = "Editor"
        4 = "Owner"
        5 = "PublishingEditor"
        6 = "PublishingAuthor"
        7 = "NonEditingAuthor"
        8 = "Contributor"
        9 = "AvailabilityOnly"
        10 = "LimitedDetails"
    }
    
    return $permissionLevels[[int]$choice]
}

# Function to grant OTD calendar permission
function Grant-OTDCalendarPermission {
    param(
        [string]$CalendarPath,
        [string]$OTDAccount,
        [string]$PermissionLevel
    )
    
    try {
        Add-MailboxFolderPermission -Identity $CalendarPath -User $OTDAccount -AccessRights $PermissionLevel -ErrorAction Stop
        return @{
            Success = $true
            Message = "Permission '$PermissionLevel' granted to $OTDAccount successfully"
        }
    }
    catch {
        # If permission already exists, try to modify it
        if ($_.Exception.Message -like "*already has permission*") {
            try {
                Set-MailboxFolderPermission -Identity $CalendarPath -User $OTDAccount -AccessRights $PermissionLevel -ErrorAction Stop
                return @{
                    Success = $true
                    Message = "Permission updated to '$PermissionLevel' for $OTDAccount successfully"
                }
            }
            catch {
                return @{
                    Success = $false
                    Message = "Failed to update permission: $($_.Exception.Message)"
                }
            }
        }
        else {
            return @{
                Success = $false
                Message = "Failed to grant permission: $($_.Exception.Message)"
            }
        }
    }
}

# Main script execution
try {
    Write-Host "Starting Exchange Mailbox Analysis..." -ForegroundColor Green
    Write-Host "CSV File: $CsvPath" -ForegroundColor Yellow
    Write-Host "Checking permissions for: $CheckUser" -ForegroundColor Yellow
    Write-Host "Output will be saved to: $OutputPath" -ForegroundColor Yellow
    Write-Host ""

    # Check if CSV file exists
    if (-not (Test-Path $CsvPath)) {
        throw "CSV file not found: $CsvPath"
    }

    # Read the CSV file
    # Assuming the CSV has a column named "Mailbox", "EmailAddress", "Identity", or "UserPrincipalName"
    $mailboxList = Import-Csv $CsvPath
    
    # Auto-detect the column name for mailbox identities
    $mailboxColumn = $null
    $possibleColumns = @("Mailbox", "EmailAddress", "Identity", "UserPrincipalName", "Email", "PrimarySmtpAddress")
    
    foreach ($col in $possibleColumns) {
        if ($mailboxList[0].PSObject.Properties.Name -contains $col) {
            $mailboxColumn = $col
            break
        }
    }
    
    if (-not $mailboxColumn) {
        Write-Host "Available columns in CSV:" -ForegroundColor Red
        $mailboxList[0].PSObject.Properties.Name | ForEach-Object { Write-Host "  - $_" }
        throw "Could not find a mailbox column. Please ensure your CSV has one of these columns: $($possibleColumns -join ', ')"
    }
    
    Write-Host "Using column '$mailboxColumn' for mailbox identities" -ForegroundColor Green
    Write-Host ""

    # Initialize results array
    $results = @()
    $totalMailboxes = $mailboxList.Count
    $currentCount = 0

    # Process each mailbox
    foreach ($row in $mailboxList) {
        $currentCount++
        $mailboxIdentity = $row.$mailboxColumn
        
        Write-Progress -Activity "Processing Mailboxes" -Status "Checking $mailboxIdentity ($currentCount of $totalMailboxes)" -PercentComplete (($currentCount / $totalMailboxes) * 100)
        
        Write-Host "Processing: $mailboxIdentity" -ForegroundColor Cyan
        
        try {
            # Check EWS status
            $casMailbox = Get-CASMailbox -Identity $mailboxIdentity -ErrorAction Stop
            $ewsEnabled = $casMailbox.EWSEnabled
            
            # Check calendar permissions
            $calendarPerms = Get-CalendarPermissions -MailboxIdentity $mailboxIdentity -UserToCheck $CheckUser
            
            # Create result object
            $result = [PSCustomObject]@{
                Mailbox = $mailboxIdentity
                EWSEnabled = $ewsEnabled
                HasCalendarPermission = $calendarPerms.HasPermission
                CalendarAccessRights = $calendarPerms.AccessRights
                CalendarFolderPath = $calendarPerms.FolderPath
                Status = "Success"
                Error = ""
            }
            
            Write-Host "  EWS Enabled: $ewsEnabled" -ForegroundColor $(if ($ewsEnabled) { "Green" } else { "Red" })
            Write-Host "  Calendar Permission: $($calendarPerms.HasPermission)" -ForegroundColor $(if ($calendarPerms.HasPermission) { "Green" } else { "Yellow" })
            if ($calendarPerms.HasPermission) {
                Write-Host "  Access Rights: $($calendarPerms.AccessRights)" -ForegroundColor Green
            }
        }
        catch {
            Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
            
            $result = [PSCustomObject]@{
                Mailbox = $mailboxIdentity
                EWSEnabled = "Error"
                HasCalendarPermission = $false
                CalendarAccessRights = "Error retrieving permissions"
                CalendarFolderPath = "N/A"
                Status = "Failed"
                Error = $_.Exception.Message
            }
        }
        
        $results += $result
        Write-Host ""
    }

    # Export results to CSV
    $results | Export-Csv -Path $OutputPath -NoTypeInformation
    
    # Display summary
    Write-Host "Analysis Complete!" -ForegroundColor Green
    Write-Host "Results saved to: $OutputPath" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Summary:" -ForegroundColor Cyan
    Write-Host "Total Mailboxes Processed: $($results.Count)"
    Write-Host "EWS Enabled: $(($results | Where-Object { $_.EWSEnabled -eq $true }).Count)"
    Write-Host "EWS Disabled: $(($results | Where-Object { $_.EWSEnabled -eq $false }).Count)"
    Write-Host "With Calendar Permissions: $(($results | Where-Object { $_.HasCalendarPermission -eq $true }).Count)"
    Write-Host "Errors: $(($results | Where-Object { $_.Status -eq 'Failed' }).Count)"
    
}
catch {
    Write-Error "Script failed: $($_.Exception.Message)"
    exit 1
}

# Example CSV format (save as mailboxes.csv):
<#
Mailbox
user1@domain.com
user2@domain.com
user3@domain.com
#>

# Usage examples:
# .\CheckMailboxEWSAndPermissions.ps1 -CsvPath "C:\temp\mailboxes.csv"
# .\CheckMailboxEWSAndPermissions.ps1 -CsvPath "C:\temp\mailboxes.csv" -CheckUser "admin@domain.com"
# .\CheckMailboxEWSAndPermissions.ps1 -CsvPath "C:\temp\mailboxes.csv" -CheckUser "otd@domain.com" -OutputPath "C:\reports\mailbox_report.csv"
