param($SCPath, $Sub, $Intag, $Resources, $Task , $File, $SmaResources, $TableStyle) 
If ($Task -eq 'Processing') {

    $FRONTDOOR = $Resources | Where-Object { $_.TYPE -eq 'microsoft.network/frontdoors' }
    $tmp = @()
    $Subs = $Sub

    foreach ($1 in $FRONTDOOR) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
            foreach ($TagKey in $Tag.Keys) {  
                $obj = @{
                    'Subscription'   = $sub1.name;
                    'Resource Group' = $1.RESOURCEGROUP;
                    'Name'           = $1.NAME;
                    'Location'       = $1.LOCATION;
                    'Friendly Name'  = $data.friendlyName;
                    'cName'          = $data.cName;
                    'State'          = $data.enabledState;
                    'Frontend'       = [string]$data.frontendEndpoints.name;
                    'Backend'        = [string]$data.backendPools.name;
                    'Health Probe'   = [string]$data.healthProbeSettings.name;
                    'Load Balancing' = [string]$data.loadBalancingSettings.name;
                    'Routing Rules'  = [string]$data.routingRules.name;
                    'Resource U'     = $ResUCount;
                    'Tag Name'       = [string]$TagKey;
                    'Tag Value'      = [string]$Tag.$TagKey
                }
                $tmp += $obj
                if ($ResUCount -eq 1) { $ResUCount = 0 } 
            }
        }
        else {    
            $obj = @{
                'Subscription'   = $sub1.name;
                'Resource Group' = $1.RESOURCEGROUP;
                'Name'           = $1.NAME;
                'Location'       = $1.LOCATION;
                'Friendly Name'  = $data.friendlyName;
                'cName'          = $data.cName;
                'State'          = $data.enabledState;
                'Frontend'       = [string]$data.frontendEndpoints.name;
                'Backend'        = [string]$data.backendPools.name;
                'Health Probe'   = [string]$data.healthProbeSettings.name;
                'Load Balancing' = [string]$data.loadBalancingSettings.name;
                'Routing Rules'  = [string]$data.routingRules.name;
                'Resource U'     = $ResUCount;
                'Tag Name'       = $null;
                'Tag Value'      = $null
            }
            $tmp += $obj
            if ($ResUCount -eq 1) { $ResUCount = 0 } 
        }
    }
    $tmp
}
Else {
    if ($SmaResources.FRONTDOOR) {
        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat 0

        $ExcelFrontDoor = $SmaResources.FrontDoor

        if ($InTag -eq $True) {
            $ExcelFrontDoor | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'Friendly Name',
            'cName',
            'State',
            'Frontend',
            'Backend',
            'Health Probe',
            'Load Balancing',
            'Routing Rules',
            'Tag Name',
            'Tag Value' | 
            Export-Excel -Path $File -WorksheetName 'FrontDoor' -AutoSize -TableName 'AzureFrontDoor' -TableStyle $tableStyle -Style $Style
        }
        else {
            $ExcelFrontDoor | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'Friendly Name',
            'cName',
            'State',
            'Frontend',
            'Backend',
            'Health Probe',
            'Load Balancing',
            'Routing Rules' | 
            Export-Excel -Path $File -WorksheetName 'FrontDoor' -AutoSize -TableName 'AzureFrontDoor' -TableStyle $tableStyle -Style $Style
        }
    }
}