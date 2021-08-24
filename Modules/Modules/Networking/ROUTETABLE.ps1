param($SCPath, $Sub, $Intag, $Resources, $Task , $File, $SmaResources, $TableStyle) 
If ($Task -eq 'Processing') {

    $ROUTETABLE = $Resources | Where-Object { $_.TYPE -eq 'microsoft.network/routetables' }

    $tmp = @()
    $Subs = $Sub

    foreach ($1 in $ROUTETABLE) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value } 
        if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
            foreach ($TagKey in $Tag.Keys) { 
                $obj = @{
                    'Subscription'                  = $sub1.name;
                    'Resource Group'                = $1.RESOURCEGROUP;
                    'Name'                          = $1.NAME;
                    'Location'                      = $1.LOCATION;
                    'Disable BGP Route Propagation' = $data.disableBgpRoutePropagation;
                    'Routes'                        = [string]$data.routes.name;
                    'Routes Prefixes'               = [string]$data.routes.properties.addressPrefix;
                    'Routes BGP Override'           = [string]$data.routes.properties.hasBgpOverride;
                    'Routes Next Hop IP'            = [string]$data.routes.properties.nextHopIpAddress;
                    'Routes Next Hop Type'          = [string]$data.routes.properties.nextHopType;
                    'Resource U'                    = $ResUCount;
                    'Tag Name'                      = [string]$TagKey;
                    'Tag Value'                     = [string]$Tag.$TagKey
                }
                $tmp += $obj
                if ($ResUCount -eq 1) { $ResUCount = 0 } 
            }
        }
        else {  
            $obj = @{
                'Subscription'                  = $sub1.name;
                'Resource Group'                = $1.RESOURCEGROUP;
                'Name'                          = $1.NAME;
                'Location'                      = $1.LOCATION;
                'Disable BGP Route Propagation' = $data.disableBgpRoutePropagation;
                'Routes'                        = [string]$data.routes.name;
                'Routes Prefixes'               = [string]$data.routes.properties.addressPrefix;
                'Routes BGP Override'           = [string]$data.routes.properties.hasBgpOverride;
                'Routes Next Hop IP'            = [string]$data.routes.properties.nextHopIpAddress;
                'Routes Next Hop Type'          = [string]$data.routes.properties.nextHopType;
                'Resource U'                    = $ResUCount;
                'Tag Name'                      = $null;
                'Tag Value'                     = $null
            }
            $tmp += $obj
            if ($ResUCount -eq 1) { $ResUCount = 0 } 
        }
    }
    $tmp
}
Else {
    if ($SmaResources.ROUTETABLE) {
        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat 0

        $ExcelRouteTable = $SmaResources.ROUTETABLE

        if ($InTag -eq $True) {
            $ExcelRouteTable | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'Disable BGP Route Propagation',
            'Routes',
            'Routes Prefixes',
            'Routes BGP Override',
            'Routes Next Hop IP',
            'Routes Next Hop Type',
            'Tag Name',
            'Tag Value' | 
            Export-Excel -Path $File -WorksheetName 'Route Tables' -AutoSize -TableName 'AzureRouteTables' -TableStyle $tableStyle -Style $Style
        }
        else {
            $ExcelRouteTable | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'Disable BGP Route Propagation',
            'Routes',
            'Routes Prefixes',
            'Routes BGP Override',
            'Routes Next Hop IP',
            'Routes Next Hop Type' | 
            Export-Excel -Path $File -WorksheetName 'Route Tables' -AutoSize -TableName 'AzureRouteTables' -TableStyle $tableStyle -Style $Style
        }
    }
}