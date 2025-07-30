# Exchange Mailbox EWS and Calendar Permissions Checker
# This script reads mailboxes from a CSV file, checks EWS status, and verifies calendar permissions

param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath,
    
    [Parameter(Mandatory=$false)]
    [string]$CheckUser = "otd@domain.com",
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "MailboxReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Function to check calendar permissions
function Get-CalendarPermissions {
    param(
        [string]$MailboxIdentity,
        [string]$UserToCheck
    )
    
    try {
        # Get the calendar folder for the mailbox
        $calendarPath = "${MailboxIdentity}:\Calendar"
        
        # Try different calendar folder names for different languages
        $calendarPaths = @(
            "${MailboxIdentity}:\Calendar",
            "${MailboxIdentity}:\Kalender",  # German
            "${MailboxIdentity}:\Calendrier", # French
            "${MailboxIdentity}:\Calendario"  # Spanish/Italian
        )
        
        $permissions = $null
        foreach ($path in $calendarPaths) {
            try {
                $permissions = Get-MailboxFolderPermission -Identity $path -User $UserToCheck -ErrorAction Stop
                break
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
            }
        }
        else {
            return @{
                HasPermission = $false
                AccessRights = "No permissions found"
                FolderPath = "N/A"
            }
        }
    }
    catch {
        return @{
            HasPermission = $false
            AccessRights = "Error: $($_.Exception.Message)"
            FolderPath = "N/A"
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
