# Init Variables
$outputPath = "C:\Intune_Reports"  # Changed to a system-accessible path
$NewFilename = "IntuneCPCDeviceInventory_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$ApplicationID = "redacted"
$TenantID = "redacted"
$AccessSecret = "redacted"
$tempBasePath = "C:\Windows\Temp"  # Using system temp folder instead of user temp

# Create output directory if it doesn't exist
if (-not (Test-Path -Path $outputPath -PathType Container)) {
    try {
        New-Item -Path $outputPath -ItemType Directory -Force | Out-Null
        Write-Host "Created output directory: $outputPath"
    } catch {
        Write-Error "Failed to create output directory: $_"
        exit 1
    }
}

# Log file setup
$logFile = Join-Path $outputPath "IntuneReport_Log_$(Get-Date -Format 'yyyyMMdd').txt"
function Write-Log {
    param (
        [string]$Message
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $logFile -Append
    Write-Host $Message
}

Write-Log "Script started"

try {
    #Create a hash table with the required value to connect to Microsoft graph
    $Body = @{    
        grant_type    = "client_credentials"
        scope         = "https://graph.microsoft.com/.default"
        client_id     = $ApplicationID
        client_secret = $AccessSecret
    } 

    #Connect to Microsoft Graph REST web service
    Write-Log "Connecting to Microsoft Graph"
    $ConnectGraph = Invoke-RestMethod -Uri https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token -Method POST -Body $Body

    #Endpoint Analytics Graph API
    $GraphGroupUrl = "https://graph.microsoft.com/beta/deviceManagement/reports/exportJobs"

    # define request body as PS Object
    $requestBody = @{
        reportName = "Devices"
        filter = "(EnrollmentType eq '10')"
        select = @(
            "DeviceId"
            "DeviceName"
            "UPN"
            "SerialNumber"
            "OS"
            "ManagedBy"
            "EnrollmentType"
            "Manufacturer"
            "Model"
        )
    }

    # Convert PS Object to JSON object
    $requestJSONBody = ConvertTo-Json $requestBody

    #define header, use the token from the above rest call to AAD
    $headers = @{
        'Authorization' = $("{0} {1}" -f $ConnectGraph.token_type,$ConnectGraph.access_token)
        'Accept' = 'application/json'
        'Content-Type' = "application/json"
    }

    #This API call will start a process in the background to download the file
    Write-Log "Starting export job"
    $webResponse = Invoke-RestMethod $GraphGroupUrl -Method 'POST' -Headers $headers -Body $requestJSONBody

    #If the call is a success, proceed to get the CSV file
    if (-not ($null -eq $webResponse)) {
        #Check status of export (GET) until status = complete
        $attempts = 0
        $maxAttempts = 30  # Maximum attempts (5 minutes with 10-second intervals)
        
        do {
            $attempts++
            #format the URL to make a next call to get the file location
            $url2GetCSV = $("https://graph.microsoft.com/beta/deviceManagement/reports/exportJobs('{0}')" -f $webResponse.id)
            Write-Log "Checking export status (Attempt $attempts of $maxAttempts): $url2GetCSV"
            $responseforCSV = Invoke-RestMethod $url2GetCSV -Method 'GET' -Headers $headers
            
            if ((-not ($null -eq $responseforCSV)) -and ($responseforCSV.status -eq "completed")) {
                try {
                    # Create temporary paths for download and extraction
                    $randomId = [System.Guid]::NewGuid().ToString("N")
                    $tempZipFile = Join-Path $tempBasePath "IntuneTempDownload_$randomId.zip"
                    $tempExtractPath = Join-Path $tempBasePath "IntuneTempExtract_$randomId"
                    
                    Write-Log "Downloading to temporary file: $tempZipFile"
                    
                    # Download to temp location
                    Invoke-WebRequest -Uri $responseforCSV.url -OutFile $tempZipFile
                    
                    Write-Log "Extracting to temporary location: $tempExtractPath"
                    
                    # Create temp extract directory
                    if (-not (Test-Path -Path $tempExtractPath)) {
                        New-Item -Path $tempExtractPath -ItemType Directory -Force | Out-Null
                    }
                    
                    # Extract to temp location
                    Expand-Archive -LiteralPath $tempZipFile -DestinationPath $tempExtractPath -Force
                    
                    # Find and copy the CSV file to final destination
                    $extractedFile = Get-ChildItem -Path $tempExtractPath -Filter "*.csv" | Select-Object -First 1
                    
                    if ($extractedFile) {
                        $finalFilePath = Join-Path $outputPath $NewFilename
                        Copy-Item -Path $extractedFile.FullName -Destination $finalFilePath -Force
                        Write-Log "File successfully extracted and saved to: $finalFilePath"
                    } else {
                        Write-Log "No CSV file found in the extraction directory."
                    }
                    
                    # Clean up temp files
                    Write-Log "Cleaning up temporary files..."
                    if (Test-Path -Path $tempZipFile) {
                        Remove-Item -Path $tempZipFile -Force
                    }
                    if (Test-Path -Path $tempExtractPath) {
                        Remove-Item -Path $tempExtractPath -Recurse -Force
                    }
                } 
                catch {
                    Write-Log "Error processing file: $_"
                }
                
                # Exit the loop once processing is complete
                break
            } 
            else {
                Write-Log "Export job still in progress..."
            }
            
            # Check if we've exceeded maximum attempts
            if ($attempts -ge $maxAttempts) {
                Write-Log "Maximum attempts reached. Export job may be stuck or taking too long."
                break
            }
            
            Start-Sleep -Seconds 10 # Delay for 10 seconds
        } While ((-not ($null -eq $responseforCSV)) -and ($responseforCSV.status -eq "inprogress"))
    } else {
        Write-Log "Failed to start export job."
    }
} catch {
    Write-Log "Error: $_"
}

Write-Log "Script completed"
