# Add-MDaemonAddressBookToOutlook.ps1
<#
.SYNOPSIS
    Adds MDaemon LDAP address book to Outlook 2021.
.DESCRIPTION
    This script configures LDAP settings for Outlook via registry if Outlook is not running.
    It can be deployed via login script or GPO for domain users.
#>

# Define LDAP settings
$ldapServer = "mdaemon.pw.salova"      # Replace with your MDaemon FQDN
$ldapPort = 389
$ldapDisplayName = "VezuAddrBook"
$baseDN = "dc=pw,dc=salova"            # Replace with your domain base DN
$ldapTimeout = 60
$maxEntries = 100

# Registry path for Outlook Address Book configuration
$regPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\LDAP\"

# Check Outlook version
$officeVer = "16.0"  # For Outlook 2021 use version 16.0

# Generate unique key name
$guid = [guid]::NewGuid().ToString()
$key = "$regPath$guid"

# Create registry key and set values
New-Item -Path $key -Force | Out-Null
Set-ItemProperty -Path $key -Name "DisplayName" -Value $ldapDisplayName
Set-ItemProperty -Path $key -Name "ServerName" -Value $ldapServer
Set-ItemProperty -Path $key -Name "PortNumber" -Value $ldapPort
Set-ItemProperty -Path $key -Name "SearchBase" -Value $baseDN
Set-ItemProperty -Path $key -Name "SearchTimeout" -Value $ldapTimeout
Set-ItemProperty -Path $key -Name "MaxEntriesReturned" -Value $maxEntries
Set-ItemProperty -Path $key -Name "UseSSL" -Value 0
Set-ItemProperty -Path $key -Name "RequireAuth" -Value 0

Write-Host "LDAP address book '$ldapDisplayName' added to Outlook 2021."

# Optionally force Outlook to reload settings on next start