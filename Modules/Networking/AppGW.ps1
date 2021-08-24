param($SCPath, $Sub, $Intag, $Resources, $Task , $File, $SmaResources, $TableStyle) 
If ($Task -eq 'Processing') {

    $APPGTW = $Resources | Where-Object { $_.TYPE -eq 'microsoft.network/applicationgateways' }

    $tmp = @()
    $Subs = $Sub

    foreach ($1 in $APPGTW) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
            foreach ($TagKey in $Tag.Keys) {        
                $obj = @{
                    'Subscription'          = $sub1.name;
                    'Resource Group'        = $1.RESOURCEGROUP;
                    'Name'                  = $1.NAME;
                    'Location'              = $1.LOCATION;
                    'State'                 = $data.OperationalState;
                    'SKU Name'              = $data.sku.tier;
                    'SKU Capacity'          = $data.sku.capacity;
                    'Backend'               = [string]$data.backendAddressPools.name;
                    'Frontend'              = [string]$data.frontendIPConfigurations.name;
                    'Frontend Ports'        = [string]$data.frontendports.properties.port;
                    'Gateways'              = [string]$data.gatewayIPConfigurations.name;
                    'HTTP Listeners'        = [string]$data.httpListeners.name;
                    'Request Routing Rules' = [string]$data.RequestRoutingRules.Name;
                    'Resource U'            = $ResUCount;
                    'Tag Name'              = [string]$TagKey;
                    'Tag Value'             = [string]$Tag.$TagKey
                }
                $tmp += $obj
                if ($ResUCount -eq 1) { $ResUCount = 0 } 
            }
        }
        else {        
            $obj = @{
                'Subscription'          = $sub1.name;
                'Resource Group'        = $1.RESOURCEGROUP;
                'Name'                  = $1.NAME;
                'Location'              = $1.LOCATION;
                'State'                 = $data.OperationalState;
                'SKU Name'              = $data.sku.tier;
                'SKU Capacity'          = $data.sku.capacity;
                'Backend'               = [string]$data.backendAddressPools.name;
                'Frontend'              = [string]$data.frontendIPConfigurations.name;
                'Frontend Ports'        = [string]$data.frontendports.properties.port;
                'Gateways'              = [string]$data.gatewayIPConfigurations.name;
                'HTTP Listeners'        = [string]$data.httpListeners.name;
                'Request Routing Rules' = [string]$data.RequestRoutingRules.Name;
                'Resource U'            = $ResUCount;
                'Tag Name'              = $null;
                'Tag Value'             = $null
            }
            $tmp += $obj
            if ($ResUCount -eq 1) { $ResUCount = 0 } 
        }
    }
    $tmp
}
Else {
    if ($SmaResources.APPGTW) {
        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat 0

        $ExcelAppGateway = $SmaResources.APPGTW

        if ($InTag -eq $True) {
            $ExcelAppGateway | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'State',
            'SKU Name',
            'SKU Capacity',
            'Backend',
            'Frontend',
            'Frontend Ports',
            'Gateways',
            'HTTP Listeners',
            'Request Routing Rules',
            'Tag Name',
            'Tag Value' | 
            Export-Excel -Path $File -WorksheetName 'App Gateway' -AutoSize -TableName 'AzureAppGateway' -TableStyle $tableStyle -Style $Style
        }
        else {
            $ExcelAppGateway | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'State',
            'SKU Name',
            'SKU Capacity',
            'Backend',
            'Frontend',
            'Frontend Ports',
            'Gateways',
            'HTTP Listeners',
            'Request Routing Rules' | 
            Export-Excel -Path $File -WorksheetName 'App Gateway' -AutoSize -TableName 'AzureAppGateway' -TableStyle $tableStyle -Style $Style
        }
    }
}