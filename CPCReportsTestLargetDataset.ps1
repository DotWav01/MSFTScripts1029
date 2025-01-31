# Cloud PC Usage Report Script
# Requires Microsoft.Graph.Beta module

# Function to format the output as a readable table
function Format-CloudPCReport {
    param (
        [Parameter(Mandatory = $true)]
        [Object[]]$Reports
    )
    
    $Reports | Format-Table -AutoSize -Property `
        @{Name = 'PC Name'; Expression = { $_.ManagedDeviceName } },
        @{Name = 'User'; Expression = { $_.UserPrincipalName } },
        @{Name = 'Usage (Hours)'; Expression = { [math]::Round($_.TotalUsageInHour, 2) } },
        @{Name = 'Last Active'; Expression = { $_.LastActiveTime } },
        @{Name = 'PC Type'; Expression = { $_.PcType } }
}

# Function to get report with specific time filter
function Get-CloudPCUsageReport {
    param (
        [Parameter(Mandatory = $true)]
        [String]$TimeFrame
    )

    $baseParams = @{
        top = 999  # Maximum allowed per request
        skip = 0
        select = @(
            "CloudPcId"
            "ManagedDeviceName"
            "UserPrincipalName"
            "TotalUsageInHour"
            "LastActiveTime"
            "PcType"
            "CreatedDate"
        )
    }

    # Calculate the date ranges
    $now = Get-Date
    switch ($TimeFrame) {
        "24hours" {
            $hours = 24
            $baseParams.filter = "(LastActiveTime gt $($now.AddHours(-$hours).ToString('yyyy-MM-ddTHH:mm:ssZ')))"
        }
        "1week" {
            $days = 7
            $baseParams.filter = "(LastActiveTime gt $($now.AddDays(-$days).ToString('yyyy-MM-ddTHH:mm:ssZ')))"
        }
        "2weeks" {
            $days = 14
            $baseParams.filter = "(LastActiveTime gt $($now.AddDays(-$days).ToString('yyyy-MM-ddTHH:mm:ssZ')))"
        }
        "4weeks" {
            $days = 28
            $baseParams.filter = "(LastActiveTime gt $($now.AddDays(-$days).ToString('yyyy-MM-ddTHH:mm:ssZ')))"
        }
        "inactive60days" {
            $days = 60
            $baseParams.filter = "(LastActiveTime lt $($now.AddDays(-$days).ToString('yyyy-MM-ddTHH:mm:ssZ')))"
        }
        default {
            Write-Error "Invalid time frame specified. Valid options: 24hours, 1week, 2weeks, 4weeks, inactive60days"
            return
        }
    }

    try {
        $allReports = @()
        $skip = 0
        do {
            $baseParams.skip = $skip
            $response = Get-MgBetaDeviceManagementVirtualEndpointReportTotalAggregatedRemoteConnectionReport -BodyParameter $baseParams
            $allReports += $response
            $skip += 999
        } while ($response.Count -eq 999)
        $reports = $allReports
        if ($reports.Count -eq 0) {
            Write-Host "No Cloud PCs found matching the specified criteria for $TimeFrame timeframe."
            return
        }
        return $reports
    }
    catch {
        Write-Error "Error retrieving Cloud PC report: $_"
        return
    }
}

# Main script
function Get-AllCloudPCReports {
    # Check if Microsoft.Graph.Beta module is installed
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Beta)) {
        Write-Error "Microsoft.Graph.Beta module is not installed. Please install it using: Install-Module Microsoft.Graph.Beta -Force"
        return
    }

    # App registration authentication parameters
    $tenantId = "<your-tenant-id>"
    $clientId = "<your-client-id>"
    $clientSecret = "<your-client-secret>"

    # Connect to Microsoft Graph using app registration
    try {
        Get-MgContext -ErrorAction Stop
    }
    catch {
        Write-Host "Connecting to Microsoft Graph..."
        $body = @{
            Grant_Type    = "client_credentials"
            Scope        = "https://graph.microsoft.com/.default"
            Client_Id    = $clientId
            Client_Secret = $clientSecret
        }
        
        Connect-MgGraph -ClientId $clientId -TenantId $tenantId -ClientSecret $clientSecret
    }

    $timeFrames = @("24hours", "1week", "2weeks", "4weeks", "inactive60days")

    foreach ($timeFrame in $timeFrames) {
        Write-Host "`n=== Cloud PC Usage Report - $timeFrame ===" -ForegroundColor Cyan
        $reports = Get-CloudPCUsageReport -TimeFrame $timeFrame
        if ($reports) {
            Format-CloudPCReport -Reports $reports
        }
    }
}

# Export functions
Export-ModuleMember -Function Get-AllCloudPCReports
