# Remove-LDAPAddressBooks.ps1

$regPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\LDAP\"
$entries = Get-ChildItem -Path $regPath

foreach ($entry in $entries) {
    $dispName = (Get-ItemProperty -Path $entry.PSPath).DisplayName
    if ($dispName -eq "MDaemon Address Book") {
        Remove-Item -Path $entry.PSPath -Recurse -Force
        Write-Host "Removed LDAP address book: $dispName"
    }
}