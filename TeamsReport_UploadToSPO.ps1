# Teams Reports Application registration details
$teamsAppId = "redacted"
$teamsTenantId = "redacted"
$teamsClientSecret = "redacted"

# SharePoint Application registration details
$sharepointAppId = "redacted"
$sharepointTenantId = "redacted"
$sharepointClientSecret = "redacted"

# SharePoint configuration
$sharepointSiteUrl = "https://yourtenant.sharepoint.com/sites/yoursite"
$sharepointSiteID = "your-site-id"  # You'll need to provide this
$sharepointLibraryURL = "https://yourtenant.sharepoint.com/sites/yoursite/Shared%20Documents"  # URL to your document library

# Note: reportFolder and logFolder are now dynamically set based on current date (see below)

# Function to get access token (Graph API style)
function Get-AccessToken {
    param(
        [string]$AppId,
        [string]$TenantId,
        [string]$ClientSecret,
        [string]$Scope
    )
    
    $body = @{
        Grant_Type = "client_credentials";
        Scope = $Scope;
        Client_Id = $AppId;
        Client_Secret = $ClientSecret
    }
    
    $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $tokenResponse = Invoke-RestMethod -Uri $tokenEndpoint -Method Post -Body $body
    return $tokenResponse.access_token
}

# Function to get SharePoint drive ID
function Get-SharePointDriveID {
    param(
        [string]$SiteID,
        [string]$LibraryURL,
        [hashtable]$Headers
    )
    
    $GraphUrl = "https://graph.microsoft.com/v1.0/sites/$SiteID/drives"
    $Result = Invoke-RestMethod -Uri $GraphUrl -Method 'GET' -Headers $Headers -ContentType "application/json"
    $DriveID = $Result.value | Where-Object {$_.webURL -eq $LibraryURL } | Select-Object id -ExpandProperty id
    
    if ($DriveID -eq $null) {
        throw "SharePoint Library under $LibraryURL could not be found."
    }
    
    return $DriveID
}

# Function to upload file to SharePoint
function Upload-FileToSharePoint {
    param(
        [string]$FilePath,
        [string]$DriveID,
        [string]$FolderPath,
        [hashtable]$Headers
    )
    
    $fileName = Split-Path $FilePath -Leaf
    $Url = "https://graph.microsoft.com/v1.0/drives/$($DriveID)/items/root:/$($FolderPath)/$($fileName):/content"
    
    try {
        $result = Invoke-RestMethod -Uri $Url -Headers $Headers -Method Put -InFile $FilePath -ContentType 'multipart/form-data'
        return $result
    }
    catch {
        # If folder doesn't exist, the upload will fail, but we'll let the calling function handle this
        throw
    }
}

# Get timestamp for filename and folder organization
$currentDate = Get-Date
$timeStamp = $currentDate.ToString("MMddyyyy_hhmm")
$monthName = $currentDate.ToString("MMMM")  # Full month name (e.g., "August")
$year = $currentDate.ToString("yyyy")

$fileName = "TeamsUserActivity-$timeStamp.csv"
$logFileName = "TeamsActivityScript-$timeStamp.log"

# Dynamic folder paths based on current month/year
$reportFolder = "Reports/Teams Activity/$year/$monthName"  # e.g., "Reports/Teams Activity/2025/August"
$logFolder = "Logs/$year/$monthName"  # e.g., "Logs/2025/August"

# Create temporary local file paths
$tempFile = "$env:TEMP\$fileName"
$logFile = "$env:TEMP\$logFileName"

# Function to write to log file
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Add-Content -Path $logFile -Value $logEntry
    
    # Also write to console with color based on level
    switch ($Level) {
        "ERROR" { Write-Host $logEntry -ForegroundColor Red }
        "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
        default { Write-Host $logEntry -ForegroundColor White }
    }
}

try {
    Write-Log "=== Teams Activity Report Script Started ===" "INFO"
    Write-Log "Report file: $fileName" "INFO"
    Write-Log "Report will be uploaded to: $reportFolder" "INFO"
    Write-Log "Log will be uploaded to: $logFolder" "INFO"
    Write-Log "Target SharePoint location: $sharepointSiteUrl" "INFO"
    
    # STEP 1: Get Teams report using Teams app credentials
    Write-Log "Getting Teams access token..." "INFO"
    $teamsAccessToken = Get-AccessToken -AppId $teamsAppId -TenantId $teamsTenantId -ClientSecret $teamsClientSecret -Scope "https://graph.microsoft.com/.default"
    $teamsSecureToken = ConvertTo-SecureString $teamsAccessToken -AsPlainText -Force
    Write-Log "Teams access token obtained successfully" "SUCCESS"
    
    # Connect to Graph with Teams credentials
    Write-Log "Connecting to Graph for Teams reports..." "INFO"
    Connect-MgGraph -AccessToken $teamsSecureToken -NoWelcome
    Write-Log "Connected to Graph with Teams credentials" "SUCCESS"
    
    # Get Teams user activity details and save to temp file
    Write-Log "Retrieving Teams user activity data for the last 7 days..." "INFO"
    $activity = Get-MgReportTeamUserActivityUserDetail -Period D7 -OutFile $tempFile
    Write-Log "Teams report saved to temporary file: $tempFile" "SUCCESS"
    
    # Disconnect from Graph
    Disconnect-MgGraph
    Write-Log "Disconnected from Graph (Teams session)" "INFO"
    
    # STEP 2: Upload to SharePoint using SharePoint app credentials and your working method
    Write-Log "Getting SharePoint access token..." "INFO"
    $sharepointAccessToken = Get-AccessToken -AppId $sharepointAppId -TenantId $sharepointTenantId -ClientSecret $sharepointClientSecret -Scope "https://graph.microsoft.com/.default"
    Write-Log "SharePoint access token obtained successfully" "SUCCESS"
    
    # Create headers for SharePoint API calls
    $Headers = @{
        Authorization = "Bearer $sharepointAccessToken";
        "Content-Type" = "application/json"
    }
    
    # Get SharePoint drive ID
    Write-Log "Getting SharePoint drive information..." "INFO"
    $DriveID = Get-SharePointDriveID -SiteID $sharepointSiteID -LibraryURL $sharepointLibraryURL -Headers $Headers
    Write-Log "SharePoint drive ID obtained: $DriveID" "SUCCESS"
    
    # Upload Teams report file
    Write-Log "Uploading Teams report file to SharePoint..." "INFO"
    try {
        $uploadResult = Upload-FileToSharePoint -FilePath $tempFile -DriveID $DriveID -FolderPath $reportFolder -Headers $Headers
        Write-Log "Successfully uploaded '$fileName' to SharePoint!" "SUCCESS"
        Write-Log "File location: $sharepointSiteUrl/$reportFolder/$fileName" "SUCCESS"
    }
    catch {
        Write-Log "Upload failed, possibly due to missing folder. Error: $($_.Exception.Message)" "WARNING"
        Write-Log "You may need to create the folder '$reportFolder' manually in SharePoint." "WARNING"
    }
    
    # Clean up temporary CSV file
    Remove-Item $tempFile -Force
    Write-Log "Temporary CSV file cleaned up" "INFO"
    
    # Upload log file
    Write-Log "Uploading log file to SharePoint..." "INFO"
    try {
        $logUploadResult = Upload-FileToSharePoint -FilePath $logFile -DriveID $DriveID -FolderPath $logFolder -Headers $Headers
        Write-Log "Log file uploaded successfully: $logFileName" "SUCCESS"
        Write-Log "Log file location: $sharepointSiteUrl/$logFolder/$logFileName" "SUCCESS"
    }
    catch {
        Write-Log "Log upload failed, possibly due to missing folder. Error: $($_.Exception.Message)" "WARNING"
        Write-Log "You may need to create the folder '$logFolder' manually in SharePoint." "WARNING"
    }
    
    Write-Log "=== Script completed successfully ===" "SUCCESS"
    
} catch {
    $errorMessage = "An error occurred: $($_.Exception.Message)"
    Write-Log $errorMessage "ERROR"
    Write-Log "Stack trace: $($_.Exception.StackTrace)" "ERROR"
    
    # Clean up temporary files if they exist
    if (Test-Path $tempFile) {
        Remove-Item $tempFile -Force
        Write-Log "Temporary CSV file cleaned up after error" "INFO"
    }
} finally {
    # Clean up temporary log file (since it's uploaded to SharePoint)
    if (Test-Path $logFile) {
        Remove-Item $logFile -Force
    }
}
