$certs = @( dir cert:\CurrentUser\my )
$certstore = new-Object System.Security.Cryptography.X509Certificates.X509Store “My”,”CurrentUser”
$certstore.Open(“ReadWrite”)

foreach ($cert in $certs) {
 write-host "Removing certificate "$cert.subject
 $certstore.Remove($cert)
}

$certstore.Close()