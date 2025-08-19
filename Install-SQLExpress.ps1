<#
.SYNOPSIS
    Installs the latest version of SQL Server Express on Windows Server 2022 Core.
.DESCRIPTION
    This script automates the download and installation of SQL Server Express.
    It performs the following actions:
    1. Checks system requirements
    2. Downloads the latest SQL Server Express installer
    3. Installs SQL Server Express with basic configuration
    4. Enables necessary firewall rules
.NOTES
    File Name      : Install-SQLExpressCore.ps1
    Prerequisites  : Windows Server 2022 Core, PowerShell 5.1 or later
    Run as Administrator
#>

#Requires -RunAsAdministrator

# Configuration parameters
$DownloadUrl = "https://go.microsoft.com/fwlink/?linkid=866658"  # SQL Server Express download link
$InstallPath = "$env:Temp\SQLExpressSetup.exe"
$InstanceName = "SQLEXPRESS"
$SqlSvcAccount = "NT AUTHORITY\NETWORK SERVICE"
$SqlSysAdmins = "BUILTIN\Administrators"

# Check if running on Server Core
$installType = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name 'InstallationType').InstallationType
if ($installType -ne "Server Core") {
    Write-Warning "This script is designed for Windows Server Core edition."
    exit 1
}

# Check if SQL Server is already installed
if (Get-Service -Name "MSSQL`$$InstanceName" -ErrorAction SilentlyContinue) {
    Write-Host "SQL Server Express ($InstanceName) is already installed." -ForegroundColor Yellow
    exit 0
}

# Download SQL Server Express
try {
    Write-Host "Downloading SQL Server Express installer..." -ForegroundColor Cyan
    Invoke-WebRequest -Uri $DownloadUrl -OutFile $InstallPath -UseBasicParsing
    Write-Host "Download completed successfully." -ForegroundColor Green
}
catch {
    Write-Error "Failed to download SQL Server Express: $_"
    exit 1
}

# Install SQL Server Express
try {
    Write-Host "Installing SQL Server Express..." -ForegroundColor Cyan
    
    $installArgs = @(
        "/QS",  # Quiet simple mode
        "/ACTION=Install",
        "/FEATURES=SQLENGINE,CONN,SDK",
        "/INSTANCENAME=$InstanceName",
        "/SQLSVCACCOUNT=`"$SqlSvcAccount`"",
        "/SQLSYSADMINACCOUNTS=`"$SqlSysAdmins`"",
        "/TCPENABLED=1",
        "/IACCEPTSQLSERVERLICENSETERMS",
        "/INDICATEPROGRESS"
    )
    
    $process = Start-Process -FilePath $InstallPath -ArgumentList $installArgs -Wait -PassThru
    
    if ($process.ExitCode -ne 0) {
        throw "Installation failed with exit code $($process.ExitCode)"
    }
    
    Write-Host "SQL Server Express installed successfully." -ForegroundColor Green
}
catch {
    Write-Error "Installation failed: $_"
    exit 1
}
finally {
    # Clean up installer
    if (Test-Path $InstallPath) {
        Remove-Item $InstallPath -Force -ErrorAction SilentlyContinue
    }
}

# Configure firewall rules
try {
    Write-Host "Configuring firewall rules..." -ForegroundColor Cyan
    New-NetFirewallRule -DisplayName "SQL Server ($InstanceName)" -Direction Inbound -LocalPort 1433 -Protocol TCP -Action Allow | Out-Null
    New-NetFirewallRule -DisplayName "SQL Server Browser" -Direction Inbound -LocalPort 1434 -Protocol UDP -Action Allow | Out-Null
    Write-Host "Firewall rules configured." -ForegroundColor Green
}
catch {
    Write-Warning "Failed to configure firewall rules: $_"
}

# Verify installation
try {
    Write-Host "Verifying installation..." -ForegroundColor Cyan
    $service = Get-Service -Name "MSSQL`$$InstanceName" -ErrorAction Stop
    if ($service.Status -ne "Running") {
        Start-Service $service -ErrorAction Stop
    }
    Write-Host "SQL Server Express is running." -ForegroundColor Green
    
    # Basic connection test
    if (Get-Command -Name "sqlcmd" -ErrorAction SilentlyContinue) {
        $query = "SELECT @@VERSION"
        $result = sqlcmd -S ".\$InstanceName" -Q $query -b -h -1
        Write-Host "SQL Server version:`n$result" -ForegroundColor Cyan
    }
}
catch {
    Write-Warning "Verification failed: $_"
}

Write-Host "SQL Server Express installation completed." -ForegroundColor Green