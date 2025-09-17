<#
.SYNOPSIS
    Creates a VPN connection with IPv6 disabled and remote gateway removed.
.DESCRIPTION
    This script creates a VPN connection, disables IPv6 for the connection,
    and removes the remote gateway option in advanced TCP/IP settings.
.NOTES
    File Name      : Create-VPNConnection.ps1
    Requires       : Run as Administrator
#>

# Parameters - modify these as needed
$VPNName = "MySecureVPN"
$ServerAddress = "vpn.example.com"
$PresharedKey = "YourSharedKeyHere"  # Leave empty if not using L2TP/IPsec with PSK
$VPNType = "L2TP"  # Can be "PPTP", "L2TP", "SSTP", or "IKEv2"
$RememberCredential = $true
$SplitTunneling = $true  # Set to $false to use remote gateway

# Run as Administrator check
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Please run this script as Administrator!"
    break
}

# Create VPN connection
try {
    Write-Host "Creating VPN connection '$VPNName'..."
    Add-VpnConnection -Name $VPNName -ServerAddress $ServerAddress -TunnelType $VPNType -RememberCredential:$RememberCredential -SplitTunneling:$SplitTunneling -ErrorAction Stop
    
    # If using L2TP with preshared key
    if ($VPNType -eq "L2TP" -and $PresharedKey) {
        Set-VpnConnectionIPsecConfiguration -ConnectionName $VPNName -AuthenticationTransformConstants GCMAES256 -CipherTransformConstants GCMAES256 -EncryptionMethod AES256 -IntegrityCheckMethod SHA384 -PfsGroup None -DHGroup Group14 -Force -PassThru
        Set-VpnConnection -Name $VPNName -SplitTunneling $SplitTunneling -RememberCredential $RememberCredential -AuthenticationMethod MSChapv2 -EncryptionLevel Required -L2tpPsk $PresharedKey -Force
    }
    
    Write-Host "VPN connection created successfully." -ForegroundColor Green
}
catch {
    Write-Host "Error creating VPN connection: $_" -ForegroundColor Red
    exit
}

# Disable IPv6 and remove remote gateway
try {
    # Get the VPN interface GUID
    $vpnInterface = Get-VpnConnection -Name $VPNName
    $interfaceGuid = $vpnInterface.InterfaceGuid
    
    # Path to the interface in the registry
    $interfaceRegPath = "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters\Interfaces\$interfaceGuid"
    
    # Disable IPv6
    Write-Host "Disabling IPv6 for VPN interface..."
    if (-not (Test-Path $interfaceRegPath)) {
        New-Item -Path $interfaceRegPath -Force | Out-Null
    }
    Set-ItemProperty -Path $interfaceRegPath -Name "DisabledComponents" -Value 0xFF -Type DWord
    
    # Remove remote gateway (even if split tunneling is enabled, sometimes Windows adds it)
    Write-Host "Removing remote gateway option..."
    $vpnInterface | Set-VpnConnection -SplitTunneling $true -Force
    
    # Additional registry tweak to ensure remote gateway is disabled
    $vpnRegPath = "HKLM:\SYSTEM\CurrentControlSet\Services\RemoteAccess\Parameters\Ikev2\"
    if (-not (Test-Path $vpnRegPath)) {
        New-Item -Path $vpnRegPath -Force | Out-Null
    }
    Set-ItemProperty -Path $vpnRegPath -Name "SkipDefaultGateway" -Value 1 -Type DWord
    
    Write-Host "IPv6 disabled and remote gateway removed successfully." -ForegroundColor Green
}
catch {
    Write-Host "Error configuring VPN settings: $_" -ForegroundColor Red
}

# Restart relevant services to apply changes
try {
    Write-Host "Restarting services to apply changes..."
    Restart-Service -Name "RemoteAccess" -Force -ErrorAction SilentlyContinue
    Restart-Service -Name "iphlpsvc" -Force -ErrorAction SilentlyContinue
}
catch {
    Write-Host "Error restarting services: $_" -ForegroundColor Yellow
}

Write-Host "VPN setup completed. You may need to reboot for all changes to take effect." -ForegroundColor Cyan