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
$targetFolderPath = "Shared Documents/Reports/Teams Activity"  # Modify this path as needed
$logFolderPath = "Shared Documents/Logs"  # Folder for log files

# Function to get access token
function Get-AccessToken {
    param(
        [string]$AppId,
        [string]$TenantId,
        [string]$ClientSecret,
        [string]$Scope
    )
    
    $body = @{
        Grant_Type = "client_credentials"
        Scope = $Scope
        Client_Id = $AppId
        Client_Secret = $ClientSecret
    }
    
    $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $tokenResponse = Invoke-RestMethod -Uri $tokenEndpoint -Method Post -Body $body
    return $tokenResponse.access_token
}

# Get timestamp for filename
$timeStamp = Get-Date -Format "MMddyyyy_hhmm"
$fileName = "TeamsUserActivity-$timeStamp.csv"
$logFileName = "TeamsActivityScript-$timeStamp.log"

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
    Write-Log "Target SharePoint location: $sharepointSiteUrl/$targetFolderPath" "INFO"
    
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
    
    # STEP 2: Upload to SharePoint using SharePoint app credentials
    Write-Log "Getting SharePoint access token..." "INFO"
    $sharepointAccessToken = Get-AccessToken -AppId $sharepointAppId -TenantId $sharepointTenantId -ClientSecret $sharepointClientSecret -Scope "https://graph.microsoft.com/.default"
    $sharepointSecureToken = ConvertTo-SecureString $sharepointAccessToken -AsPlainText -Force
    Write-Log "SharePoint access token obtained successfully" "SUCCESS"
    
    # Connect to Graph with SharePoint credentials
    Write-Log "Connecting to Graph for SharePoint operations..." "INFO"
    Connect-MgGraph -AccessToken $sharepointSecureToken -NoWelcome
    Write-Log "Connected to Graph with SharePoint credentials" "SUCCESS"
    
    # Get SharePoint site ID
    Write-Log "Getting SharePoint site information..." "INFO"
    $siteId = (Get-MgSite -Search (Split-Path $sharepointSiteUrl -Leaf)).Id
    Write-Log "SharePoint site ID obtained: $siteId" "SUCCESS"
    
    # Get the drive (document library) ID
    $drives = Get-MgSiteDrive -SiteId $siteId
    $driveId = ($drives | Where-Object { $_.Name -eq "Documents" }).Id
    Write-Log "Document library ID obtained: $driveId" "SUCCESS"
    
# Function to create folder structure and return parent ID
function Create-SharePointFolder {
    param(
        [string]$DriveId,
        [string]$FolderPath
    )
    
    $folderParts = $FolderPath.Split('/')
    $parentId = "root"
    
    foreach ($folderPart in $folderParts) {
        if ($folderPart -ne "Shared Documents") {  # Skip the root folder name
            try {
                # Try to get existing folder
                $folder = Get-MgDriveItem -DriveId $DriveId -DriveItemId $parentId -ChildName $folderPart -ErrorAction SilentlyContinue
                if (-not $folder) {
                    # Create folder if it doesn't exist
                    $newFolder = @{
                        name = $folderPart
                        folder = @{}
                    }
                    $folder = New-MgDriveItem -DriveId $DriveId -ParentId $parentId -BodyParameter $newFolder
                    Write-Log "Created folder: $folderPart" "SUCCESS"
                } else {
                    Write-Log "Folder already exists: $folderPart" "INFO"
                }
                $parentId = $folder.Id
            }
            catch {
                Write-Log "Failed to create/access folder '$folderPart': $($_.Exception.Message)" "ERROR"
                throw
            }
        }
    }
    return $parentId
}
    
    # Create folder structure for reports
    Write-Log "Ensuring report folder structure exists: $targetFolderPath" "INFO"
    $reportParentId = Create-SharePointFolder -DriveId $driveId -FolderPath $targetFolderPath
    
    # Upload file to SharePoint
    Write-Log "Uploading file to SharePoint..." "INFO"
    $fileContent = Get-Content $tempFile -Raw -Encoding UTF8
    $fileBytes = [System.Text.Encoding]::UTF8.GetBytes($fileContent)
    Write-Log "File size: $($fileBytes.Length) bytes" "INFO"
    
    $uploadParams = @{
        DriveId = $driveId
        ParentId = $reportParentId
        Name = $fileName
        ContentBytes = $fileBytes
    }
    
    $uploadedFile = Set-MgDriveItemContent @uploadParams
    Write-Log "Successfully uploaded '$fileName' to SharePoint!" "SUCCESS"
    Write-Log "File location: $sharepointSiteUrl/$targetFolderPath/$fileName" "SUCCESS"
    
    # Clean up temporary CSV file
    Remove-Item $tempFile -Force
    Write-Log "Temporary CSV file cleaned up" "INFO"
    
    # Create folder structure for logs
    Write-Log "Ensuring log folder structure exists: $logFolderPath" "INFO"
    $logParentId = Create-SharePointFolder -DriveId $driveId -FolderPath $logFolderPath
    
    # Upload log file to SharePoint in logs folder
    Write-Log "Uploading log file to SharePoint logs folder..." "INFO"
    $logContent = Get-Content $logFile -Raw -Encoding UTF8
    $logBytes = [System.Text.Encoding]::UTF8.GetBytes($logContent)
    
    $logUploadParams = @{
        DriveId = $driveId
        ParentId = $logParentId
        Name = $logFileName
        ContentBytes = $logBytes
    }
    
    $uploadedLogFile = Set-MgDriveItemContent @logUploadParams
    Write-Log "Log file uploaded successfully: $logFileName" "SUCCESS"
    Write-Log "Log file location: $sharepointSiteUrl/$logFolderPath/$logFileName" "SUCCESS"
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
    
    # Ensure we disconnect from Graph
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-Log "Disconnected from Graph" "INFO"
    } catch {
        # Ignore errors during disconnect
    }
