# Set execution policy for this session only
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

# Remediation Script - Runs as SYSTEM, blocks logged-in user access
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
    # STEP 1: Hide the folder FIRST
    $folder = Get-Item $recordingsPath -Force
    $folder.Attributes = $folder.Attributes -bor [System.IO.FileAttributes]::Hidden
    
    # STEP 2: Get the user's SID
    $userAccount = New-Object System.Security.Principal.NTAccount($loggedInUser)
    $userSID = $userAccount.Translate([System.Security.Principal.SecurityIdentifier])
    
    # STEP 3: Modify the ACL
    $acl = Get-Acl $recordingsPath
    
    # Remove any existing rules for this user (to avoid duplicates)
    $acl.Access | Where-Object { $_.IdentityReference.Value -eq $userSID.Value } | ForEach-Object {
        $acl.RemoveAccessRule($_) | Out-Null
    }
    
    # Ensure SYSTEM has full control (Allow rule)
    $systemSID = New-Object System.Security.Principal.SecurityIdentifier("S-1-5-18")
    $systemRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
        $systemSID,
        "FullControl",
        "ContainerInherit,ObjectInherit",
        "None",
        "Allow"
    )
    $acl.AddAccessRule($systemRule)
    
    # Add deny rule for the logged-in user
    $denyRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
        $userSID, 
        "FullControl", 
        "ContainerInherit,ObjectInherit", 
        "None", 
        "Deny"
    )
    $acl.AddAccessRule($denyRule)
    
    # Apply the ACL
    Set-Acl $recordingsPath $acl
    
    Write-Output "Successfully configured Recordings folder for $username - User blocked, SYSTEM has access"
    Exit 0
} catch {
    Write-Output "Error during remediation: $_"
    Exit 1
}
