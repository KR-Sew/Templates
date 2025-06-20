Get-NetFirewallRule | Where-Object { $_.ElementName -like "*1c*" } | ForEach-Object {
    $rule = $_
    $addressFilters = Get-NetFirewallAddressFilter -AssociatedNetFirewallRule $rule
    $portFilters = Get-NetFirewallPortFilter -AssociatedNetFirewallRule $rule

    # Create a custom object to hold the information
    [PSCustomObject]@{
        RuleName      = $rule.DisplayName
        LocalAddress  = $addressFilters.LocalAddress -join ', '
        RemoteAddress = $addressFilters.RemoteAddress -join ', '
        LocalPort     = $portFilters.LocalPort -join ', '
        RemotePort    = $portFilters.RemotePort -join ', '
    }
} | Format-List
