$ProgressPreference='SilentlyContinue'
$base="https://openvpn.net/community-downloads/"
$link=((Invoke-WebRequest -Uri "$base" -UseBasicParsing).links.href | Select-String -Pattern ".*amd64.msi$" | Select-Object -First 1 | out-string).Trim("")

Write-Host "Downloading the latest OpenVPN MSI for AMD64"
Invoke-WebRequest -Uri "$link" -Outfile $env:temp\openvpn.msi

Write-Host "Installing OpenVPN"
Start-Process "$env:temp\openvpn.msi" -ArgumentList "/qn" -Wait
Write-Host "INSTALLED!"

Write-Host "Excluding OpenVPN from firewall"
netsh advfirewall firewall add rule name="OpenVPN CLI" dir=in action=allow program="C:\Program Files\OpenVPN\bin\openvpn.exe" enable=yes profile=domain,public,private
netsh advfirewall firewall add rule name="OpenVPN GUI" dir=in action=allow program="C:\Program Files\OpenVPN\bin\openvpn-gui.exe" enable=yes profile=domain,public,private
Write-Host "DONE!"

Write-Host "Creating OpenVPN Administrators Group"
New-LocalGroup "OpenVPN Administrators"
(Get-LocalGroupMember users).name | foreach {
  Write-Host "Adding $_ to OpenVPN Administrators group"
  Add-LocalGroupmember "OpenVPN Administrators" $_
}
