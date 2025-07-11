$processName = "TimeControl"
$process = Get-Process -Name $processName -ErrorAction SilentlyContinue

if ($process) {
    $pidId = $process.Id

    # Monitor TCP Connections
    $tcpConnections = Get-NetTCPConnection | Where-Object { $_.OwningProcess -eq $pidId }
    $tcpConnections | Format-Table -Property LocalAddress, LocalPort, RemoteAddress, RemotePort, State

    # Monitor UDP Endpoints
    $udpEndpoints = Get-NetUDPEndpoint | Where-Object { $_.OwningProcess -eq $pidId }
    $udpEndpoints | Format-Table -Property LocalAddress, LocalPort
} else {
    Write-Host "Process '$processName' not found."
}