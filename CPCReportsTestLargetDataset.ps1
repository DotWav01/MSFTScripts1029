# App registration parameters
$tenantId = "<your-tenant-id>"
$clientId = "<your-client-id>"
$clientSecret = "<your-client-secret>"

# Connect to Microsoft Graph
Connect-MgGraph -ClientId $clientId -TenantId $tenantId -ClientSecret $clientSecret

# Create output directory
$outputPath = "C:\CloudPCReports"
New-Item -ItemType Directory -Path $outputPath -Force

# Time frames to check
$timeFrames = @("24hours", "1week", "2weeks", "4weeks", "inactive60days")

foreach ($timeFrame in $timeFrames) {
    $now = Get-Date
    $baseParams = @{
        top = 999
        skip = 0
        select = @(
            "ManagedDeviceName"
            "UserPrincipalName"
            "TotalUsageInHour"
            "LastActiveTime"
            "PcType"
        )
    }

    # Set filter based on timeframe
    switch ($timeFrame) {
        "24hours" { $baseParams.filter = "(LastActiveTime gt $($now.AddHours(-24).ToString('yyyy-MM-ddTHH:mm:ssZ')))" }
        "1week" { $baseParams.filter = "(LastActiveTime gt $($now.AddDays(-7).ToString('yyyy-MM-ddTHH:mm:ssZ')))" }
        "2weeks" { $baseParams.filter = "(LastActiveTime gt $($now.AddDays(-14).ToString('yyyy-MM-ddTHH:mm:ssZ')))" }
        "4weeks" { $baseParams.filter = "(LastActiveTime gt $($now.AddDays(-28).ToString('yyyy-MM-ddTHH:mm:ssZ')))" }
        "inactive60days" { $baseParams.filter = "(LastActiveTime lt $($now.AddDays(-60).ToString('yyyy-MM-ddTHH:mm:ssZ')))" }
    }

    $allReports = @()
    $skip = 0
    do {
        $baseParams.skip = $skip
        $response = Get-MgBetaDeviceManagementVirtualEndpointReportTotalAggregatedRemoteConnectionReport -BodyParameter $baseParams
        $allReports += $response
        $skip += 999
    } while ($response.Count -eq 999)

    # Export to CSV
    $fileName = "CloudPC_Report_$($timeFrame)_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $filePath = Join-Path $outputPath $fileName
    $allReports | Select-Object ManagedDeviceName, UserPrincipalName, TotalUsageInHour, LastActiveTime, PcType | 
        Export-Csv -Path $filePath -NoTypeInformation
    Write-Host "Report exported to: $filePath"
}
