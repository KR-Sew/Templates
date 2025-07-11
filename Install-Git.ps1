<#
.SYNOPSIS
    Downloads and installs the latest Git for Windows
.DESCRIPTION
    This script checks GitHub for the latest Git for Windows release,
    downloads the installer, and performs a silent installation.
.NOTES
    File Name      : Install-LatestGit.ps1
    Prerequisite   : PowerShell 5.1 or later
#>

# Set execution policy to allow script execution (if needed)
# Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force

# Function to get the latest Git for Windows release version
function Get-LatestGitVersion {
    try {
        $apiUrl = "https://api.github.com/repos/git-for-windows/git/releases/latest"
        $response = Invoke-RestMethod -Uri $apiUrl -UseBasicParsing
        $version = $response.tag_name -replace '^v', ''  # Remove 'v' prefix if present
        
        # Find the 64-bit standalone installer
        $asset = $response.assets | Where-Object {
            $_.name -match '^Git-\d+\.\d+\.\d+-64-bit\.exe$'
        } | Select-Object -First 1
        
        if (-not $asset) {
            throw "Could not find 64-bit installer in release assets"
        }

        return @{
            Version = $version
            DownloadUrl = $asset.browser_download_url
            FileName = $asset.name
        }
    }
    catch {
        Write-Host "Failed to get latest Git version: $_" -ForegroundColor Red
        exit 1
    }
}

# Get latest version info
Write-Host "Checking for latest Git for Windows version..."
$latestGit = Get-LatestGitVersion
Write-Host "Latest version found: $($latestGit.Version)" -ForegroundColor Cyan

# Download the installer
$installerPath = "$env:TEMP\$($latestGit.FileName)"
try {
    Write-Host "Downloading Git $($latestGit.Version) installer..."
    Invoke-WebRequest -Uri $latestGit.DownloadUrl -OutFile $installerPath -UseBasicParsing
    Write-Host "Download completed." -ForegroundColor Green
}
catch {
    Write-Host "Failed to download Git installer: $_" -ForegroundColor Red
    exit 1
}

# Install Git silently with common options
$installArgs = @(
    "/VERYSILENT",          # No progress or prompts
    "/SUPPRESSMSGBOXES",    # No message boxes
    "/NORESTART",           # Don't prompt for restart
    "/NOCANCEL",            # Don't allow user to cancel
    "/SP-",                 # Don't show "This will install..." prompt
    "/LOG",                 # Create a log file
    "/COMPONENTS=icons,ext\reg\shellhere,assoc,assoc_sh", # Standard components
    "/D=C:\Program Files\Git" # Install location
)

try {
    Write-Host "Installing Git $($latestGit.Version)..."
    Start-Process -FilePath $installerPath -ArgumentList $installArgs -Wait
    Write-Host "Git installation completed successfully." -ForegroundColor Green
    
    # Add Git to PATH if not already there
    $envPath = [Environment]::GetEnvironmentVariable("Path", "Machine")
    if ($envPath -notlike "*Git\cmd*") {
        $gitPath = "C:\Program Files\Git\cmd"
        [Environment]::SetEnvironmentVariable("Path", "$envPath;$gitPath", "Machine")
        Write-Host "Added Git to system PATH." -ForegroundColor Green
    }
    
    # Clean up installer
    Remove-Item -Path $installerPath -Force
}
catch {
    Write-Host "Git installation failed: $_" -ForegroundColor Red
    exit 1
}

# Verify installation
try {
    $gitVersion = git --version 2>$null
    if ($gitVersion) {
        Write-Host "Git installed successfully: $gitVersion" -ForegroundColor Green
        Write-Host "You may need to restart your terminal or computer for PATH changes to take effect." -ForegroundColor Yellow
    }
    else {
        Write-Host "Git installation verification failed." -ForegroundColor Yellow
    }
}
catch {
    Write-Host "Git installation verification failed. You may need to restart your shell or computer." -ForegroundColor Yellow
}