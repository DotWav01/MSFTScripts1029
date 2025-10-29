# Set execution policy for this session only
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

# Detection Script - Runs as SYSTEM, detects logged-in user
# Exit 0 = Compliant (no action needed)
# Exit 1 = Non-Compliant (remediation needed)

# Get the currently logged-in user
$loggedInUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
if (-not $loggedInUser) {
    Write-Output "No user currently logged in"
    Exit 0
}

# Extract just the username (remove domain)
$username = $loggedInUser.Split('\')[-1]

# Build the path to the user's OneDrive Recordings folder
$userProfile = "C:\Users\$username"
$recordingsPath = "$userProfile\OneDrive\Recordings"

if (-not (Test-Path $recordingsPath)) {
    Write-Output "Recordings folder does not exist for $username"
    Exit 0
}

try {
    # Check if folder is hidden
    $folder = Get-Item $recordingsPath -Force
    $isHidden = $folder.Attributes -band [System.IO.FileAttributes]::Hidden
    
    # Get the user's SID
    $userAccount = New-Object System.Security.Principal.NTAccount($loggedInUser)
    $userSID = $userAccount.Translate([System.Security.Principal.SecurityIdentifier])
    
    # Check if access is denied for the user
    $acl = Get-Acl $recordingsPath
    $denyRule = $acl.Access | Where-Object { 
        $_.IdentityReference.Value -eq $userSID.Value -and 
        $_.AccessControlType -eq "Deny" -and
        $_.FileSystemRights -match "FullControl"
    }
    
    # Check if SYSTEM has full control
    $systemRule = $acl.Access | Where-Object {
        $_.IdentityReference.Value -eq "NT AUTHORITY\SYSTEM" -and
        $_.AccessControlType -eq "Allow" -and
        $_.FileSystemRights -match "FullControl"
    }
    
    if ($isHidden -and $denyRule -and $systemRule) {
        Write-Output "Recordings folder is properly configured for $username"
        Exit 0
    } else {
        Write-Output "Recordings folder needs remediation for $username"
        Exit 1
    }
} catch {
    Write-Output "Error checking folder status: $_"
    Exit 1
}
