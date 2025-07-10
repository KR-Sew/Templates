Import-Module DFSR

# Parameters
$replicationGroupName = "Retailing"
$newMemberServer = "V22"
$destinationPath = "R:\Shares"
$foldersToReplicate = @(
    "Директор направления E-commerce",
    "Командировки",
    "Рукводитель Фулфилмент",
    "Руководитель МНОГОРУК",
    "Руководитель Сервисы для ритейла",
    "Сервисы для ритейла"
)

# Check if the destination path exists, create if it doesn't
if (-not (Test-Path -Path $destinationPath)) {
    Write-Host "Creating destination path: $destinationPath..."
    New-Item -Path $destinationPath -ItemType Directory -Force | Out-Null
}

# Add server to the replication group
try {
    Add-DfsrMember -GroupName $replicationGroupName -ComputerName $newMemberServer
    Write-Host "Server $newMemberServer successfully added to replication group $replicationGroupName."
} catch {
    Write-Warning "Error adding server to replication group: $_"
}

# Configure replication for each folder
foreach ($folder in $foldersToReplicate) {
    $fullDestinationPath = Join-Path -Path $destinationPath -ChildPath $folder
    
    # Create subfolder if it doesn't exist
    if (-not (Test-Path -Path $fullDestinationPath)) {
        New-Item -Path $fullDestinationPath -ItemType Directory -Force | Out-Null
    }
    
    try {
    Set-DfsrMembership -GroupName $replicationGroupName -FolderName $folder `
                      -ComputerName $newMemberServer -ContentPath $fullDestinationPath `
                      -PrimaryMember $false -Force
    Write-Host "Replication successfully configured for folder: $folder -> $fullDestinationPath."
} catch {
    Write-Error "Error configuring replication for folder ${folder}: $_"
 }
}    

# Check results
Write-Host ""
#Write-Host "Checking replication settings for server ${newMemberServer} :"
Get-DfsrMember -GroupName $replicationGroupName -ComputerName $newMemberServer | Format-Table -AutoSize

Write-Host ""
Write-Host "Checking folder memberships:"
Get-DfsrMembership -GroupName $replicationGroupName -ComputerName $newMemberServer | 
    Select-Object FolderName, ContentPath, Enabled, State | Format-Table -AutoSize
