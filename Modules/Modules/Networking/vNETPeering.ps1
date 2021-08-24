param($SCPath, $Sub, $Intag, $Resources, $Task , $File, $SmaResources, $TableStyle) 
If ($Task -eq 'Processing') {

    $VNETPeering = $Resources | Where-Object { $_.TYPE -eq 'microsoft.network/virtualnetworks' -and $null -ne $AzNetwork.Peering -and $AzNetwork.Peering -ne '' }

    $tmp = @()
    $Subs = $Sub

    foreach ($1 in $VNETPeering) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        foreach ($2 in $data.addressSpace.addressPrefixes) {
            foreach ($4 in $data.virtualNetworkPeerings) {
                foreach ($5 in $4.properties.remoteAddressSpace.addressPrefixes) {
                    if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
                        foreach ($TagKey in $Tag.Keys) {  
                            $obj = @{
                                'Subscription'                          = $sub1.name;
                                'Resource Group'                        = $1.RESOURCEGROUP;
                                'VNET Name'                             = $1.NAME;
                                'Location'                              = $1.LOCATION;
                                'Zone'                                  = $1.ZONES;
                                'Address Space'                         = $2;
                                'Peering Name'                          = $4.name;
                                'Peering VNet'                          = $4.properties.remoteVirtualNetwork.id.split('/')[8];
                                'Peering State'                         = $4.properties.peeringState;
                                'Peering Use Remote Gateways'           = $4.properties.useRemoteGateways;
                                'Peering Allow Gateway Transit'         = $4.properties.allowGatewayTransit;
                                'Peering Allow Forwarded Traffic'       = $4.properties.allowForwardedTraffic;
                                'Peering Do Not Verify Remote Gateways' = $4.properties.doNotVerifyRemoteGateways;
                                'Peering Allow Virtual Network Access'  = $4.properties.allowVirtualNetworkAccess;
                                'Peering Address Space'                 = $5;
                                'Resource U'                            = $ResUCount;
                                'Tag Name'                              = [string]$TagKey;
                                'Tag Value'                             = [string]$Tag.$TagKey
                            }
                            $tmp += $obj
                            if ($ResUCount -eq 1) { $ResUCount = 0 } 
                        }
                    }
                    else {   
                        $obj = @{
                            'Subscription'                          = $sub1.name;
                            'Resource Group'                        = $1.RESOURCEGROUP;
                            'VNET Name'                             = $1.NAME;
                            'Location'                              = $1.LOCATION;
                            'Zone'                                  = $1.ZONES;
                            'Address Space'                         = $2;
                            'Peering Name'                          = $4.name;
                            'Peering VNet'                          = $4.properties.remoteVirtualNetwork.id.split('/')[8];
                            'Peering State'                         = $4.properties.peeringState;
                            'Peering Use Remote Gateways'           = $4.properties.useRemoteGateways;
                            'Peering Allow Gateway Transit'         = $4.properties.allowGatewayTransit;
                            'Peering Allow Forwarded Traffic'       = $4.properties.allowForwardedTraffic;
                            'Peering Do Not Verify Remote Gateways' = $4.properties.doNotVerifyRemoteGateways;
                            'Peering Allow Virtual Network Access'  = $4.properties.allowVirtualNetworkAccess;
                            'Peering Address Space'                 = $5;
                            'Resource U'                            = $ResUCount;
                            'Tag Name'                              = $null;
                            'Tag Value'                             = $null
                        }
                        $tmp += $obj
                        if ($ResUCount -eq 1) { $ResUCount = 0 } 
                    }
                }
            }
        }
            
    }
    $tmp
}
Else {
    if ($SmaResources.VNETPeering) {
        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat 0

        $ExcelPeering = $SmaResources.VNETPeering

        if ($InTag -eq $True) {
            $ExcelPeering | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Location',
            'Zone',
            'Peering Name',
            'VNET Name',
            'Address Space',
            'Peering VNet',
            'Peering Address Space',
            'Peering State',
            'Peering Use Remote Gateways',
            'Peering Allow Gateway Transit',
            'Peering Allow Forwarded Traffic',
            'Peering Do Not Verify Remote Gateways',
            'Peering Allow Virtual NetworkAccess',
            'Tag Name',
            'Tag Value' | 
            Export-Excel -Path $File -WorksheetName 'Peering' -AutoSize -TableName 'AzureVNETPeerings' -TableStyle $tableStyle -Style $Style
        }
        else {
            $ExcelPeering | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Location',
            'Zone',
            'Peering Name',
            'VNET Name',
            'Address Space',
            'Peering VNet',
            'Peering Address Space',
            'Peering State',
            'Peering Use Remote Gateways',
            'Peering Allow Gateway Transit',
            'Peering Allow Forwarded Traffic',
            'Peering Do Not Verify Remote Gateways',
            'Peering Allow Virtual NetworkAccess' | 
            Export-Excel -Path $File -WorksheetName 'Peering' -AutoSize -TableName 'AzureVNETPeerings' -TableStyle $tableStyle -Style $Style
        }

    }
}