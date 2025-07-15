# Create a self-signed code signing certificate with email and 20-year validity

$cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=IISLogsToExcel Application, E=kumar.anirudha@gmail.com" -CertStoreLocation "Cert:\CurrentUser\My" -NotAfter (Get-Date).AddYears(20)

$pwd = ConvertTo-SecureString -String "#IISLogsToExcel#" -Force -AsPlainText

Export-PfxCertificate -Cert $cert -FilePath "IISLogsToExcel.pfx" -Password $pwd

# "C:\Program Files (x86)\Microsoft SDKs\ClickOnce\SignTool\signtool.exe" sign /f "IISLogsToExcel.pfx" /p "#IISLogsToExcel#" /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 "$(ProjectDir)bin\$(ConfigurationName)\net8.0-windows\IISLogsToExcel.exe"