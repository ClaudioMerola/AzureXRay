param($SCPath, $Sub, $Intag, $Resources, $Task , $File, $SmaResources, $TableStyle) 
If ($Task -eq 'Processing') {

    $VNETGTW = $Resources | Where-Object { $_.TYPE -eq 'microsoft.network/virtualnetworkgateways' }

    $tmp = @()
    $Subs = $Sub

    foreach ($1 in $VNETGTW) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
            foreach ($TagKey in $Tag.Keys) {  
                $obj = @{
                    'Subscription'           = $sub1.name;
                    'Resource Group'         = $1.RESOURCEGROUP;
                    'Name'                   = $1.NAME;
                    'Location'               = $1.LOCATION;
                    'SKU'                    = $data.sku.tier;
                    'Active-active mode'     = $data.activeActive; 
                    'Gateway Type'           = $data.gatewayType;
                    'Gateway Generation'     = $data.vpnGatewayGeneration;
                    'VPN Type'               = $data.vpnType;
                    'Enable Private Address' = $data.enablePrivateIpAddress;
                    'Enable BGP'             = $data.enableBgp;
                    'BGP ASN'                = $data.bgpsettings.asn;
                    'BGP Peering Address'    = $data.bgpSettings.bgpPeeringAddress;
                    'BGP Peer Weight'        = $data.bgpSettings.peerWeight;
                    'Gateway Public IP'      = [string]$data.ipConfigurations.properties.publicIPAddress.id.split("/")[8];
                    'Gateway Subnet Name'    = [string]$data.ipConfigurations.properties.subnet.id.split("/")[8];
                    'Resource U'             = $ResUCount;
                    'Tag Name'               = [string]$TagKey;
                    'Tag Value'              = [string]$Tag.$TagKey
                }
                $tmp += $obj
                if ($ResUCount -eq 1) { $ResUCount = 0 } 
            }
        }
        else {
            $obj = @{
                'Subscription'           = $sub1.name;
                'Resource Group'         = $1.RESOURCEGROUP;
                'Name'                   = $1.NAME;
                'Location'               = $1.LOCATION;
                'SKU'                    = $data.sku.tier;
                'Active-active mode'     = $data.activeActive; 
                'Gateway Type'           = $data.gatewayType;
                'Gateway Generation'     = $data.vpnGatewayGeneration;
                'VPN Type'               = $data.vpnType;
                'Enable Private Address' = $data.enablePrivateIpAddress;
                'Enable BGP'             = $data.enableBgp;
                'BGP ASN'                = $data.bgpsettings.asn;
                'BGP Peering Address'    = $data.bgpSettings.bgpPeeringAddress;
                'BGP Peer Weight'        = $data.bgpSettings.peerWeight;
                'Gateway Public IP'      = [string]$data.ipConfigurations.properties.publicIPAddress.id.split("/")[8];
                'Gateway Subnet Name'    = [string]$data.ipConfigurations.properties.subnet.id.split("/")[8];
                'Resource U'             = $ResUCount;
                'Tag Name'               = $null;
                'Tag Value'              = $null
            }
            $tmp += $obj
            if ($ResUCount -eq 1) { $ResUCount = 0 } 
        }
    }
    $tmp
}
Else {
    if ($SmaResources.VNETGTW) {
        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat 0

        $ExcelVNETGTW = $SmaResources.VNETGTW         

        if ($InTag -eq $True) {
            $ExcelVNETGTW | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'SKU',
            'Active-active mode',
            'Gateway Type',
            'Gateway Generation',
            'VPN Type',
            'Enable Private Address',
            'Enable BGP',
            'BGP ASN',
            'BGP Peering Address',
            'BGP Peer Weight',
            'Gateway Public IP',
            'Gateway Subnet Name',
            'Tag Name',
            'Tag Value'  | 
            Export-Excel -Path $File -WorksheetName 'VNET Gateways' -AutoSize -TableName 'AzureVNETGateways' -TableStyle $tableStyle -ConditionalText $txtvnet -Style $Style
        }
        else {
            $ExcelVNETGTW | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'SKU',
            'Active-active mode',
            'Gateway Type',
            'Gateway Generation',
            'VPN Type',
            'Enable Private Address',
            'Enable BGP',
            'BGP ASN',
            'BGP Peering Address',
            'BGP Peer Weight',
            'Gateway Public IP',
            'Gateway Subnet Name' | 
            Export-Excel -Path $File -WorksheetName 'VNET Gateways' -AutoSize -TableName 'AzureVNETGateways' -TableStyle $tableStyle -ConditionalText $txtvnet -Style $Style
        }
    }
}