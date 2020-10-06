$store = "cert:\CurrentUser\My"

$paramsAlice = @{
 CertStoreLocation = $store
 Subject = "CN=Alice"
 KeyLength = 8192
 KeyAlgorithm = "RSA" 
 KeyUsage = "DataEncipherment"
 Type = "DocumentEncryptionCert"
}

$paramsBob = @{
 CertStoreLocation = $store
 Subject = "CN=Arno"
 KeyLength = 4096
 KeyAlgorithm = "RSA" 
 KeyUsage = "DataEncipherment"
 Type = "DocumentEncryptionCert"
}

# generate new certificate and add it to certificate store
$certAlice = New-SelfSignedCertificate @paramsAlice
$certBob = New-SelfSignedCertificate @paramsBob

$bytes = [byte[]]($certAliceBytes)
$Endian = if([System.BitConverter]::IsLittleEndian){1,0}else{0,1};$bytes=[byte[]]($certAliceBytes[$Endian]);
[bitconverter]::ToInt16($bytes,0)
2014

$certBobBytes = $certBob.GetPublicKey()
[bitconverter]::ToInt32($certBobBytes,0)

# list all certs 
Get-ChildItem -path $store
pause
# Encryption / Decryption

$message = "My secret message"

$cipher = $message  | Protect-CmsMessage -To "CN=Bob" 
Write-Host "Cipher:" -ForegroundColor Green
$cipher

Write-Host "Decrypted message:" -ForegroundColor Green
$cipher | Unprotect-CmsMessage


# Exporting/Importing certificate

$pwd = ("P@ssword" | ConvertTo-SecureString -AsPlainText -Force)
$privateKey = "$home\Documents\Test1.pfx"
$publicKey = "$home\Documents\Test1.cer"

# Export private key as PFX certificate, to use those Keys on different machine/user
Export-PfxCertificate -FilePath $privateKey -Cert $cert -Password $pwd

# Export Public key, to share with other users
Export-Certificate -FilePath $publicKey -Cert $cert

#Remove certificate from store
$cert | Remove-Item

# Add them back:
# Add private key on your machine
Import-PfxCertificate -FilePath $privateKey -CertStoreLocation $store -Password $pwd

# This is for other users (so they can send you encrypted messages)
Import-Certificate -FilePath $publicKey -CertStoreLocation $store