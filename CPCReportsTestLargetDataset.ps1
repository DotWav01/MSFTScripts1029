# Import required modules
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Graph.DeviceManagement -Scope CurrentUser
Install-Module ImportExcel -Scope CurrentUser

# Authentication parameters
$tenantId = "YOUR_TENANT_ID"
$clientId = "YOUR_CLIENT_ID"
$clientSecret = "YOUR_CLIENT_SECRET"

# Function to connect to Microsoft Graph
function Connect-ToGraph {
    try {
        $tokenBody = @{
            Grant_Type    = "client_credentials"
            Scope        = "https://graph.microsoft.com/.default"
            Client_Id    = $clientId
            Client_Secret = $clientSecret
        }
        
        $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method POST -Body $tokenBody
        $headers = @{
            "Authorization" = "Bearer $($tokenResponse.access_token)"
            "Content-Type"  = "application/json"
        }
        return $headers
    }
    catch {
        Write-Error "Failed to connect to Microsoft Graph: $_"
        exit 1
    }
}

# Function to get Cloud PC usage data with filter
function Get-CloudPCUsage {
    param (
        [Parameter(Mandatory=$true)]
        [hashtable]$Headers,
        
        [Parameter(Mandatory=$true)]
        [int]$DaysFilter
    )
    
    # Format date in ISO 8601 format with timezone
    $filterDate = [System.Web.HttpUtility]::UrlEncode((Get-Date).AddDays(-$DaysFilter).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ"))
    
    $filter = if ($DaysFilter -eq 60) {
        "`$filter=lastConnectDateTime lt ${filterDate}"
    } else {
        "`$filter=lastConnectDateTime ge ${filterDate}"
    }
    
    try {
        $allResults = @()
        # Add System.Web for URL encoding
        Add-Type -AssemblyName System.Web

        $baseUri = "https://graph.microsoft.com/v1.0/deviceManagement/virtualEndpoint/reports/getCloudPCConnectivityHistory"
        $uri = "${baseUri}?${filter}&`$top=999"
        
        Write-Host "Request URI: $uri" -ForegroundColor Yellow  # Debug output
        
        do {
            $response = Invoke-RestMethod -Uri $uri -Headers $Headers -Method GET -TimeoutSec 120
            
            if ($response.value) {
                $allResults += $response.value
                Write-Host "Retrieved $($allResults.Count) records so far..."
            }
            
            $uri = $response.'@odata.nextLink'
        } while ($uri)
        
        return $allResults
    }
    catch {
        Write-Error "Failed to fetch Cloud PC usage data: $_"
        return $null
    }
}

# Function to create Excel report
function Create-ExcelReport {
    param (
        [Parameter(Mandatory=$true)]
        [hashtable]$Data
    )
    
    $excelPath = "CloudPC_Usage_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
    
    try {
        $excel = Open-ExcelPackage -Path $excelPath -Create
        
        foreach ($period in $Data.Keys) {
            Write-Host "Creating worksheet for $period..."
            $reportData = $Data[$period] | Select-Object `
                @{N='Device Name';E={$_.cloudPcName}}, `
                @{N='User Principal Name';E={$_.userPrincipalName}}, `
                @{N='Date Created';E={$_.createDateTime}}, `
                @{N='Last Connection Date';E={$_.lastConnectDateTime}}, `
                @{N='Connection Duration (hours)';E={[math]::Round($_.connectionDuration/3600, 2)}}, `
                @{N='Status';E={$_.status}}
            
            $reportData | Export-Excel -ExcelPackage $excel -WorksheetName $period -AutoSize -TableName $period -BoldTopRow
            
            # Add filters and freeze panes
            $worksheet = $excel.Workbook.Worksheets[$period]
            $worksheet.View.FreezePanes(2, 1)
        }
        
        Close-ExcelPackage $excel
        Write-Host "Report generated successfully: $excelPath"
    }
    catch {
        Write-Error "Failed to create Excel report: $_"
    }
}

# Main script execution
try {
    # Connect to Graph
    $headers = Connect-ToGraph

    # Define time periods for filtering
    $timePeriods = @{
        "Last_24_Hours" = 1
        "Last_Week" = 7
        "Last_2_Weeks" = 14
        "Last_4_Weeks" = 28
        "No_Connection_60_Days" = 60
    }

    # Collect data for each time period
    $allData = @{}
    foreach ($period in $timePeriods.Keys) {
        Write-Host "Fetching data for $period..."
        $data = Get-CloudPCUsage -Headers $headers -DaysFilter $timePeriods[$period]
        if ($data) {
            $allData[$period] = $data
        }
    }

    # Generate Excel report
    if ($allData.Count -gt 0) {
        Create-ExcelReport -Data $allData
    }
    else {
        Write-Error "No data collected for any time period"
    }
}
catch {
    Write-Error "Script execution failed: $_"
}
