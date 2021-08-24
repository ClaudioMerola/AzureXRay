param($SCPath, $Sub, $Intag, $Resources, $Task , $File, $SmaResources, $TableStyle) 
If ($Task -eq 'Processing') {

    $VNET = $Resources | Where-Object { $_.TYPE -eq 'microsoft.network/virtualnetworks' }

    $tmp = @()
    $Subs = $Sub

    foreach ($1 in $VNET) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        foreach ($2 in $data.addressSpace.addressPrefixes) {
            foreach ($3 in $data.subnets) {
                if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
                    foreach ($TagKey in $Tag.Keys) {
                        $obj = @{
                            'Subscription'                                 = $sub1.name;
                            'Resource Group'                               = $1.RESOURCEGROUP;
                            'Name'                                         = $1.NAME;
                            'Location'                                     = $1.LOCATION;
                            'Zone'                                         = $1.ZONES;
                            'Address Space'                                = $2;
                            'Enable DDOS Protection'                       = $data.enableDdosProtection;
                            'Enable VM Protection'                         = $data.enableVmProtection;
                            'Subnet Name'                                  = $3.name;
                            'Subnet Prefix'                                = $3.properties.addressPrefix;
                            'Subnet Private Link Service Network Policies' = $3.properties.privateLinkServiceNetworkPolicies;
                            'Subnet Private Endpoint Network Policies'     = $3.properties.privateEndpointNetworkPolicies;
                            'Resource U'                                   = $ResUCount;
                            'Tag Name'                                     = [string]$TagKey;
                            'Tag Value'                                    = [string]$Tag.$TagKey
                        }
                        $tmp += $obj
                        if ($ResUCount -eq 1) { $ResUCount = 0 } 
                    }
                }
                else {
                    $obj = @{
                        'Subscription'                                 = $sub1.name;
                        'Resource Group'                               = $1.RESOURCEGROUP;
                        'Name'                                         = $1.NAME;
                        'Location'                                     = $1.LOCATION;
                        'Zone'                                         = $1.ZONES;
                        'Address Space'                                = $2;
                        'Enable DDOS Protection'                       = $data.enableDdosProtection;
                        'Enable VM Protection'                         = $data.enableVmProtection;
                        'Subnet Name'                                  = $3.name;
                        'Subnet Prefix'                                = $3.properties.addressPrefix;
                        'Subnet Private Link Service Network Policies' = $3.properties.privateLinkServiceNetworkPolicies;
                        'Subnet Private Endpoint Network Policies'     = $3.properties.privateEndpointNetworkPolicies;
                        'Resource U'                                   = $ResUCount;
                        'Tag Name'                                     = $null;
                        'Tag Value'                                    = $null
                    }
                    $tmp += $obj
                    if ($ResUCount -eq 1) { $ResUCount = 0 } 
                }
            }
        }
    }
    $tmp
}
Else {
    if ($SmaResources.VNET) {
        $txtvnet = $(New-ConditionalText false -Range G:H
            New-ConditionalText falso -Range G:H)

        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat '0'

        $ExcelVNET = $SmaResources.VNET          

        if ($InTag -eq $True) {
            $ExcelVNET | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'Zone',
            'Address Space',
            'Enable DDOS Protection',
            'Enable VM Protection',
            'Subnet Name',
            'Subnet Prefix',
            'Subnet Private Link Service Network Policies',
            'Subnet Private Endpoint Network Policies',
            'Tag Name',
            'Tag Value'  | 
            Export-Excel -Path $File -WorksheetName 'VNET' -AutoSize -TableName 'AzureVNETs' -TableStyle $tableStyle -ConditionalText $txtvnet -Style $Style
        }
        else {
            $ExcelVNET | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'Zone',
            'Address Space',
            'Enable DDOS Protection',
            'Enable VM Protection',
            'Subnet Name',
            'Subnet Prefix',
            'Subnet Private Link Service Network Policies',
            'Subnet Private Endpoint Network Policies' | 
            Export-Excel -Path $File -WorksheetName 'VNET' -AutoSize -TableName 'AzureVNETs' -TableStyle $tableStyle -ConditionalText $txtvnet -Style $Style
        }

        $excel = Open-ExcelPackage -Path $File -KillExcel

        $null = $excel.VNET.Cells["G1"].AddComment("Azure DDoS Protection Standard, combined with application design best practices, provides enhanced DDoS mitigation features to defend against DDoS attacks.", "Azure Resource Inventory")
        $excel.VNET.Cells["G1"].Hyperlink = 'https://docs.microsoft.com/en-us/azure/ddos-protection/ddos-protection-overview'

        Close-ExcelPackage $excel 

    }
}