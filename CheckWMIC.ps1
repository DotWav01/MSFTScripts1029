# PowerShell script to check if WMIC is enabled and available

Write-Host "=== Checking WMIC Availability ===" -ForegroundColor Cyan
Write-Host ""

# Method 1: Check if wmic.exe exists
Write-Host "1. Checking if wmic.exe exists in System32:" -ForegroundColor Yellow
$wmicPath = "$env:SystemRoot\System32\wbem\wmic.exe"
if (Test-Path $wmicPath) {
    Write-Host "   ✓ WMIC.exe found at: $wmicPath" -ForegroundColor Green
    $wmicVersion = (Get-Item $wmicPath).VersionInfo.FileVersion
    Write-Host "   Version: $wmicVersion" -ForegroundColor Gray
} else {
    Write-Host "   ✗ WMIC.exe not found" -ForegroundColor Red
}
Write-Host ""

# Method 2: Try to execute WMIC command
Write-Host "2. Testing WMIC execution:" -ForegroundColor Yellow
try {
    $result = & wmic.exe os get Caption /value 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "   ✓ WMIC is executable and working" -ForegroundColor Green
        Write-Host "   Sample output: $($result[0..50] -join '')" -ForegroundColor Gray
    } else {
        Write-Host "   ✗ WMIC command failed with exit code: $LASTEXITCODE" -ForegroundColor Red
    }
} catch {
    Write-Host "   ✗ WMIC is not available or disabled" -ForegroundColor Red
    Write-Host "   Error: $_" -ForegroundColor Red
}
Write-Host ""

# Method 3: Check Windows optional features (Windows 10/11)
Write-Host "3. Checking Windows Optional Features:" -ForegroundColor Yellow
try {
    $wmicFeature = Get-WindowsOptionalFeature -Online -FeatureName "MicrosoftWindowsWMICore*" -ErrorAction SilentlyContinue
    if ($wmicFeature) {
        foreach ($feature in $wmicFeature) {
            Write-Host "   Feature: $($feature.FeatureName)" -ForegroundColor Gray
            Write-Host "   State: $($feature.State)" -ForegroundColor $(if($feature.State -eq "Enabled"){"Green"}else{"Red"})
        }
    } else {
        Write-Host "   Unable to query Windows Optional Features" -ForegroundColor Gray
    }
} catch {
    Write-Host "   Note: Optional Features check requires elevated privileges" -ForegroundColor Gray
}
Write-Host ""

# Method 4: Check if WMI service is running
Write-Host "4. Checking WMI Service Status:" -ForegroundColor Yellow
$wmiService = Get-Service -Name "Winmgmt" -ErrorAction SilentlyContinue
if ($wmiService) {
    Write-Host "   Service Name: Windows Management Instrumentation" -ForegroundColor Gray
    Write-Host "   Status: $($wmiService.Status)" -ForegroundColor $(if($wmiService.Status -eq "Running"){"Green"}else{"Red"})
    Write-Host "   Startup Type: $($wmiService.StartType)" -ForegroundColor Gray
} else {
    Write-Host "   ✗ WMI Service not found" -ForegroundColor Red
}
Write-Host ""

# Method 5: Test WMI functionality using PowerShell cmdlets
Write-Host "5. Testing WMI through PowerShell (alternative to WMIC):" -ForegroundColor Yellow
try {
    $osInfo = Get-WmiObject -Class Win32_OperatingSystem -ErrorAction Stop
    Write-Host "   ✓ WMI is functional through PowerShell" -ForegroundColor Green
    Write-Host "   OS: $($osInfo.Caption)" -ForegroundColor Gray
} catch {
    Write-Host "   ✗ WMI queries through PowerShell failed" -ForegroundColor Red
}
Write-Host ""

# Summary
Write-Host "=== Summary ===" -ForegroundColor Cyan
if ((Test-Path $wmicPath) -and ($wmiService.Status -eq "Running")) {
    Write-Host "WMIC appears to be available on your system" -ForegroundColor Green
    Write-Host ""
    Write-Host "Note: Starting with Windows 11 21H2 and Windows Server 2022," -ForegroundColor Yellow
    Write-Host "WMIC is deprecated. Use PowerShell cmdlets like Get-WmiObject" -ForegroundColor Yellow
    Write-Host "or Get-CimInstance instead." -ForegroundColor Yellow
} else {
    Write-Host "WMIC may not be available or properly configured" -ForegroundColor Red
    Write-Host "You can use PowerShell WMI cmdlets as alternatives." -ForegroundColor Yellow
}
