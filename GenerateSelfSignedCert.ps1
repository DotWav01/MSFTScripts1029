# Generate a self-signed certificate
$certName = ""
$cert = New-SelfSignedCertificate -Subject "CN=$certName" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -KeyAlgorithm RSA `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears(2)

# Export the certificate (public key) for Azure
$certPath = "C:\Temp\$certName.cer"
Export-Certificate -Cert $cert -FilePath $certPath

# Export the private key (PFX) for the script
$pfxPath = "C:\temp\$certName.pfx"
$pfxPassword = ConvertTo-SecureString -String "" -Force -AsPlainText
Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $pfxPassword

Write-Host "`nCertificate created successfully!" -ForegroundColor Green
Write-Host "Thumbprint: $($cert.Thumbprint)" -ForegroundColor Cyan
Write-Host "CER file (for Azure): $certPath" -ForegroundColor Cyan
Write-Host "PFX file (for script): $pfxPath" -ForegroundColor Cyan
