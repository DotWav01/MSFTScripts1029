# Cloud PC Usage Report Script
# Function to get report with specific time filter
function Get-CloudPCUsageReport {
    param (
        [Parameter(Mandatory = $true)]
        [String]$TimeFrame,
        [Parameter(Mandatory = $true)]
        [String]$OutputPath
    )

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
        Connect-MgGraph -ClientId $clientId -TenantId $tenantId -ClientSecret $clientSecret
    }

    $baseParams = @{
        top = 999
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
            Write-Error "Invalid time frame specified"
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

        if ($allReports.Count -eq 0) {
            Write-Host "No Cloud PCs found for $TimeFrame timeframe."
            return
        }

        # Export to CSV
        $fileName = "CloudPC_Report_$($TimeFrame)_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $filePath = Join-Path $OutputPath $fileName
        $allReports | Select-Object ManagedDeviceName, UserPrincipalName, TotalUsageInHour, LastActiveTime, PcType | 
            Export-Csv -Path $filePath -NoTypeInformation
        Write-Host "Report exported to: $filePath"
    }
    catch {
        Write-Error "Error retrieving Cloud PC report: $_"
    }
}

function Get-AllCloudPCReports {
    param(
        [Parameter(Mandatory=$true)]
        [string]$OutputPath
    )

    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force
    }

    $timeFrames = @("24hours", "1week", "2weeks", "4weeks", "inactive60days")
    foreach ($timeFrame in $timeFrames) {
        Get-CloudPCUsageReport -TimeFrame $timeFrame -OutputPath $OutputPath
    }
}
