<#
.SYNOPSIS
    Adds an LDAP address book to Microsoft Outlook (2016, 2019, 2021)
.DESCRIPTION
    This script configures an LDAP address book in Outlook by modifying registry settings.
    It creates a flag file to prevent duplicate configurations.
#>

# LDAP Configuration
$LDAPdisplayname = "VezuAddrBook"
$LDAPserver = "pw.salova"
$LDAPport = "389"
$LDAPsearchbase = "dc=pw,dc=salova"

# Determine Outlook version
function Get-OutlookVersion {
    $versions = @{
        14 = "2010"
        15 = "2013"
        16 = "2016"
        17 = "2019"
        21 = "2021"
    }

    foreach ($ver in $versions.Keys) {
        $path = "HKCU:\Software\Microsoft\Office\$ver.0\Outlook"
        if (Test-Path $path) {
            return @{
                Version = $versions[$ver]
                Number = $ver
            }
        }
    }
    return $null
}

$outlookInfo = Get-OutlookVersion

if (-not $outlookInfo) {
    Write-Host "Outlook not found or unsupported version."
    exit
}

$outlookVer = $outlookInfo.Version
$outlookNum = $outlookInfo.Number

# Set flag file path based on Outlook version
$flagFile = Join-Path $env:APPDATA "addbook_dont_remove$outlookVer.txt"

# Check if flag file exists (already configured)
if (Test-Path $flagFile) {
    Write-Host "Address book already configured for Outlook $outlookVer."
    exit
}

# Configure registry based on Outlook version
if ($outlookNum -ge 16) {
    # Outlook 2016, 2019, 2021 use the same registry path
    $registryFolder = "Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\"
} elseif ($outlookNum -eq 15) {
    # Outlook 2013
    $registryFolder = "Software\Microsoft\Office\15.0\Outlook\Profiles\Outlook\"
} elseif ($outlookNum -eq 14) {
    # Outlook 2010
    $registryFolder = "Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Outlook\"
}

# Create LDAP Type Key
$sKeyPath = $registryFolder + "e8cb48869c395445ade13e3c1c80d154\"
if (-not (Test-Path "HKCU:\$sKeyPath")) {
    New-Item -Path "HKCU:\$sKeyPath" -Force | Out-Null
}

# Set LDAP Type values
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "00033009" -Value ([byte[]](0,0,0,0)) -Type Binary
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "00033e03" -Value ([byte[]](0x23,0,0,0)) -Type Binary
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e3001" -Value "Microsoft LDAP Directory"
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e3006" -Value "Microsoft LDAP Directory"
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e300a" -Value "EMABLT.DLL"
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e3d09" -Value "EMABLT"
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e3d13" -Value "{6485D268-C2AC-11D1-AD3E-10A0C911C9C0}"
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "01023d0c" -Value ([byte[]](0x5c,0xb9,0x3b,0x24,0xff,0x71,0x07,0x41,0xb7,0xd8,0x3b,0x9c,0xb6,0x31,0x79,0x92)) -Type Binary

# Create LDAP Connection Settings Key
$sKeyPath = $registryFolder + "5cb93b24ff710741b7d83b9cb6317992\"
if (-not (Test-Path "HKCU:\$sKeyPath")) {
    New-Item -Path "HKCU:\$sKeyPath" -Force | Out-Null
}

# Set LDAP Connection values
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "00033009" -Value ([byte[]](0x20,0,0,0)) -Type Binary
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "000b6613" -Value ([byte[]](0,0)) -Type Binary
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "000b6615" -Value ([byte[]](0x01,0x00)) -Type Binary
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "000b6622" -Value ([byte[]](0x01,0x00)) -Type Binary
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e3001" -Value $LDAPdisplayname
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e3d09" -Value "EMABLT"
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e3d0a" -Value "BJABLR.DLL"
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e3d0b" -Value "ServiceEntry"
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e3d13" -Value "{6485D268-C2AC-11D1-AD3E-10A0C911C9C0}"
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e6600" -Value $LDAPserver
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e6601" -Value $LDAPport
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e6602" -Value ""
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e6603" -Value $LDAPsearchbase
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e6604" -Value "(&(mail=*)(|(mail=%s*)(|(cn=%s*)(|(sn=%s*)(givenName=%s*)))))"
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e6605" -Value "SMTP"
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e6606" -Value "mail"
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e6607" -Value "60"
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e6608" -Value "100"
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e6609" -Value "120"
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e660a" -Value "15"
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e660b" -Value ""
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e660c" -Value "OFF"
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e660d" -Value "OFF"
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e660e" -Value "NONE"
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e660f" -Value "OFF"
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e6610" -Value "postalAddress"
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e6611" -Value "cn"
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e6612" -Value "1"
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "001e67f1" -Value ([byte[]](0x0a)) -Type Binary
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "01023615" -Value ([byte[]](0x50,0xa7,0x0a,0x61,0x55,0xde,0xd3,0x11,0x9d,0x60,0x00,0xc0,0x4f,0x4c,0x8e,0xfa)) -Type Binary
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "01023d01" -Value ([byte[]](0xe8,0xcb,0x48,0x86,0x9c,0x39,0x54,0x45,0xad,0xe1,0x3e,0x3c,0x1c,0x80,0xd1,0x54)) -Type Binary
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "01026631" -Value ([byte[]](0x98,0x17,0x82,0x92,0x5b,0x43,0x03,0x4b,0x99,0x5d,0x5c,0xc6,0x74,0x88,0x7b,0x34)) -Type Binary
Set-ItemProperty -Path "HKCU:\$sKeyPath" -Name "101e3d0f" -Value ([byte[]](0x02,0x00,0x00,0x00,0x0c,0x00,0x00,0x00,0x17,0x00,0x00,0x00,0x45,0x4d,0x41,0x42,0x4c,0x54,0x2e,0x44,0x4c,0x4c,0x00,0x42,0x4a,0x41,0x42,0x4c,0x52,0x2e,0x44,0xc,0x4c,0x00)) -Type Binary

# Update Backup Key for LDAP types
$backupKeyPath = $registryFolder + "9207f3e0a3b11019908b08002b2a56c2\"
$currentValue = (Get-ItemProperty -Path "HKCU:\$backupKeyPath" -Name "01023d01")."01023d01"
$newValue = $currentValue + [byte[]](0xe8,0xcb,0x48,0x86,0x9c,0x39,0x54,0x45,0xad,0xe1,0x3e,0x3c,0x1c,0x80,0xd1,0x54)
Set-ItemProperty -Path "HKCU:\$backupKeyPath" -Name "01023d01" -Value $newValue -Type Binary

# Update Backup Key for LDAP connection settings
$currentValue = (Get-ItemProperty -Path "HKCU:\$backupKeyPath" -Name "01023d0e")."01023d0e"
$newValue = $currentValue + [byte[]](0x5c,0xb9,0x3b,0x24,0xff,0x71,0x07,0x41,0xb7,0xd8,0x3b,0x9c,0xb6,0x31,0x79,0x92)
Set-ItemProperty -Path "HKCU:\$backupKeyPath" -Name "01023d0e" -Value $newValue -Type Binary

# Delete Active Books List Key
$deleteKeyPath = $registryFolder + "9375CFF0413111d3B88A00104B2A6676\{ED475419-B0D6-11D2-8C3B-00104B2A6676}"
if (Test-Path "HKCU:\$deleteKeyPath") {
    Remove-Item -Path "HKCU:\$deleteKeyPath" -Recurse -Force
}

# Create flag file
Set-Content -Path $flagFile -Value "This file is the flag required to check that address book already installed from gpo."

Write-Host "LDAP address book successfully configured for Outlook $outlookVer."