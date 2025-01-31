# App registration parameters
$tenantId = "<your-tenant-id>"
$clientId = "<your-client-id>"
$clientSecret = "<your-client-secret>"

try {
    # Connect to Microsoft Graph with explicit scope
    Connect-MgGraph -ClientId $clientId -TenantId $tenantId -ClientSecret $clientSecret -Scopes "CloudPC.Read.All"
    
    # Test connection
    $testConnection = Get-MgContext
    Write-Host "Connected as: $($testConnection.Account)"

    # Create a simple test query
    $testParams = @{
        top = 1
        select = @("ManagedDeviceName")
    }
    
    # Test the API call
    $testCall = Get-MgBetaDeviceManagementVirtualEndpointReportTotalAggregatedRemoteConnectionReport -BodyParameter $testParams
    
    if ($testCall) {
        Write-Host "API access confirmed"
    }
} catch {
    Write-Error "Connection error: $($_.Exception.Message)"
    exit
}

$outputPath = "C:\CloudPCReports"
New-Item -ItemType Directory -Path $outputPath -Force -ErrorAction SilentlyContinue

$timeFrames = @("24hours", "1week", "2weeks", "4weeks", "inactive60days")

foreach ($timeFrame in $timeFrames) {
    $now = Get-Date
    $baseParams = @{
        top = 999
        skip = 0
        select = @(
            "ManagedDeviceName",
            "UserPrincipalName",
            "TotalUsageInHour",
            "LastActiveTime",
            "PcType"
        )
    }

    switch ($timeFrame) {
        "24hours" { $baseParams.filter = "(LastActiveTime gt $($now.AddHours(-24).ToString('yyyy-MM-ddTHH:mm:ssZ')))" }
        "1week" { $baseParams.filter = "(LastActiveTime gt $($now.AddDays(-7).ToString('yyyy-MM-ddTHH:mm:ssZ')))" }
        "2weeks" { $baseParams.filter = "(LastActiveTime gt $($now.AddDays(-14).ToString('yyyy-MM-ddTHH:mm:ssZ')))" }
        "4weeks" { $baseParams.filter = "(LastActiveTime gt $($now.AddDays(-28).ToString('yyyy-MM-ddTHH:mm:ssZ')))" }
        "inactive60days" { $baseParams.filter = "(LastActiveTime lt $($now.AddDays(-60).ToString('yyyy-MM-ddTHH:mm:ssZ')))" }
    }

    try {
        Write-Host "Fetching $timeFrame report..."
        $allReports = @()
        $skip = 0
        
        do {
            $baseParams.skip = $skip
            $response = Get-MgBetaDeviceManagementVirtualEndpointReportTotalAggregatedRemoteConnectionReport -BodyParameter $baseParams
            if ($response) {
                $allReports += $response
            }
            $skip += 999
        } while ($response -and $response.Count -eq 999)

        $fileName = "CloudPC_Report_$($timeFrame)_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $filePath = Join-Path $outputPath $fileName
        
        if ($allReports.Count -gt 0) {
            $allReports | Select-Object ManagedDeviceName, UserPrincipalName, TotalUsageInHour, LastActiveTime, PcType | 
                Export-Csv -Path $filePath -NoTypeInformation
            Write-Host "Exported $($allReports.Count) records to: $filePath"
        } else {
            Write-Host "No data found for $timeFrame"
        }
    } catch {
        Write-Error "Error processing $timeFrame report: $($_.Exception.Message)"
    }
}
