param($SCPath, $Sub, $Intag, $Resources, $Task , $File, $SmaResources, $TableStyle) 
If ($Task -eq 'Processing') {

    $LoadBalancer = $Resources | Where-Object { $_.TYPE -eq 'microsoft.network/loadbalancers' }

    $tmp = @()
    $Subs = $Sub

    foreach ($1 in $LoadBalancer) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        if ($null -ne $data.frontendIPConfigurations -and $null -ne $data.backendAddressPools -and $null -ne $data.probes) {
            foreach ($2 in $data.frontendIPConfigurations) {
                $Fronttarget = ''    
                $Frontsub = ''
                $FrontType = ''
                if ($null -ne $2.properties.subnet.id) {
                    $Fronttarget = $2.properties.subnet.id.split('/')[8]
                    $Frontsub = $2.properties.subnet.id.split('/')[10]
                    $FrontType = 'VNET' 
                }
                elseif ($null -ne $2.properties.publicIPAddress.id) {
                    $Fronttarget = $2.properties.publicIPAddress.id.split('/')[8]
                    $Frontsub = ''
                    $FrontType = 'Public IP' 
                }       
                foreach ($3 in $data.backendAddressPools) {
                    $BackTarget = ''
                    $BackType = ''
                    if ($null -ne $3.properties.backendIPConfigurations.id) {
                        $BackTarget = $3.properties.backendIPConfigurations.id.split('/')[8]
                        $BackType = $3.properties.backendIPConfigurations.id.split('/')[7]
                    }
                    foreach ($4 in $data.probes) {
                        if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
                            foreach ($TagKey in $Tag.Keys) {
                                $obj = @{
                                    'Subscription'              = $sub1.name;
                                    'Resource Group'            = $1.RESOURCEGROUP;
                                    'Name'                      = $1.NAME;
                                    'Location'                  = $1.LOCATION;
                                    'SKU'                       = $1.sku.name;
                                    'Frontend Name'             = $2.name;
                                    'Frontend Target'           = $Fronttarget;
                                    'Frontend Type'             = $FrontType;
                                    'Frontend Subnet'           = $frontsub;
                                    'Backend Pool Name'         = $3.name;
                                    'Backend Target'            = $BackTarget;
                                    'Backend Type'              = $BackType;
                                    'Probe Name'                = $4.name;
                                    'Probe Interval (sec)'      = $4.properties.intervalInSeconds;
                                    'Probe Protocol'            = $4.properties.protocol;
                                    'Probe Port'                = $4.properties.port;
                                    'Probe Unhealthy threshold' = $4.properties.numberOfProbes;
                                    'Resource U'                = $ResUCount;
                                    'Tag Name'                  = [string]$TagKey;
                                    'Tag Value'                 = [string]$Tag.$TagKey
                                }
                                $tmp += $obj
                                if ($ResUCount -eq 1) { $ResUCount = 0 } 
                            }
                        }
                        else { 
                            $obj = @{
                                'Subscription'              = $sub1.name;
                                'Resource Group'            = $1.RESOURCEGROUP;
                                'Name'                      = $1.NAME;
                                'Location'                  = $1.LOCATION;
                                'SKU'                       = $1.sku.name;
                                'Frontend Name'             = $2.name;
                                'Frontend Target'           = $Fronttarget;
                                'Frontend Type'             = $FrontType;
                                'Frontend Subnet'           = $frontsub;
                                'Backend Pool Name'         = $3.name;
                                'Backend Target'            = $BackTarget;
                                'Backend Type'              = $BackType;
                                'Probe Name'                = $4.name;
                                'Probe Interval (sec)'      = $4.properties.intervalInSeconds;
                                'Probe Protocol'            = $4.properties.protocol;
                                'Probe Port'                = $4.properties.port;
                                'Probe Unhealthy threshold' = $4.properties.numberOfProbes;
                                'Resource U'                = $ResUCount;
                                'Tag Name'                  = $null;
                                'Tag Value'                 = $null
                            }
                            $tmp += $obj
                            if ($ResUCount -eq 1) { $ResUCount = 0 } 
                        }    
                    }
                }
            }
        }  
        elseif ($null -ne $data.frontendIPConfigurations -and $null -ne $data.backendAddressPools -and $null -eq $data.probes) {
            foreach ($2 in $data.frontendIPConfigurations) {
                $Fronttarget = ''    
                $Frontsub = ''
                if ($null -ne $2.properties.subnet.id) {
                    $Fronttarget = $2.properties.subnet.id.split('/')[8]
                    $Frontsub = $2.properties.subnet.id.split('/')[10]
                    $FrontType = 'VNET' 
                }
                elseif ($null -ne $2.properties.publicIPAddress.id) {
                    $Fronttarget = $2.properties.publicIPAddress.id.split('/')[8]
                    $Frontsub = ''
                    $FrontType = 'Public IP' 
                }        
                foreach ($3 in $data.backendAddressPools) {
                    $BackTarget = ''
                    $BackType = ''
                    if ($null -ne $3.properties.backendIPConfigurations.id) {
                        $BackTarget = $3.properties.backendIPConfigurations.id.split('/')[8]
                        $BackType = $3.properties.backendIPConfigurations.id.split('/')[7]
                    }
                    if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
                        foreach ($TagKey in $Tag.Keys) {  
                            $obj = @{
                                'Subscription'              = $sub1.name;
                                'Resource Group'            = $1.RESOURCEGROUP;
                                'Name'                      = $1.NAME;
                                'Location'                  = $1.LOCATION;
                                'SKU'                       = $1.sku.name;
                                'Frontend Name'             = $2.name;
                                'Frontend Target'           = $Fronttarget;
                                'Frontend Type'             = $FrontType;
                                'Frontend Subnet'           = $frontsub;
                                'Backend Pool Name'         = $3.name;
                                'Backend Target'            = $BackTarget;
                                'Backend Type'              = $BackType;
                                'Probe Name'                = $null;
                                'Probe Interval (sec)'      = $null;
                                'Probe Protocol'            = $null;
                                'Probe Port'                = $null;
                                'Probe Unhealthy threshold' = $null;
                                'Resource U'                = $ResUCount;
                                'Tag Name'                  = [string]$TagKey;
                                'Tag Value'                 = [string]$Tag.$TagKey
                            }
                            $tmp += $obj
                            if ($ResUCount -eq 1) { $ResUCount = 0 }          
                        }
                    }
                    else {
                        $obj = @{
                            'Subscription'              = $sub1.name;
                            'Resource Group'            = $1.RESOURCEGROUP;
                            'Name'                      = $1.NAME;
                            'Location'                  = $1.LOCATION;
                            'SKU'                       = $1.sku.name;
                            'Frontend Name'             = $2.name;
                            'Frontend Target'           = $Fronttarget;
                            'Frontend Type'             = $FrontType;
                            'Frontend Subnet'           = $frontsub;
                            'Backend Pool Name'         = $3.name;
                            'Backend Target'            = $BackTarget;
                            'Backend Type'              = $BackType;
                            'Probe Name'                = $null;
                            'Probe Interval (sec)'      = $null;
                            'Probe Protocol'            = $null;
                            'Probe Port'                = $null;
                            'Probe Unhealthy threshold' = $null;
                            'Resource U'                = $ResUCount;
                            'Tag Name'                  = $null;
                            'Tag Value'                 = $null
                        }
                        $tmp += $obj 
                        if ($ResUCount -eq 1) { $ResUCount = 0 } 
                    }
                }
            }
        }   
        elseif ($null -ne $data.frontendIPConfigurations -and $null -eq $data.backendAddressPools -and $null -eq $data.probes) {
            foreach ($2 in $data.frontendIPConfigurations) {
                $Fronttarget = ''    
                $Frontsub = ''
                if ($null -ne $2.properties.subnet.id) {
                    $Fronttarget = $2.properties.subnet.id.split('/')[8]
                    $Frontsub = $2.properties.subnet.id.split('/')[10]
                    $FrontType = 'VNET' 
                }
                elseif ($null -ne $2.properties.publicIPAddress.id) {
                    $Fronttarget = $2.properties.publicIPAddress.id.split('/')[8]
                    $Frontsub = ''
                    $FrontType = 'Public IP' 
                }         
                if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
                    foreach ($TagKey in $Tag.Keys) {
                        $obj = @{
                            'Subscription'              = $sub1.name;
                            'Resource Group'            = $1.RESOURCEGROUP;
                            'Name'                      = $1.NAME;
                            'Location'                  = $1.LOCATION;
                            'SKU'                       = $1.sku.name;
                            'Frontend Name'             = $2.name;
                            'Frontend Target'           = $Fronttarget;
                            'Frontend Type'             = $FrontType;
                            'Frontend Subnet'           = $frontsub;
                            'Backend Pool Name'         = $null;
                            'Backend Target'            = $null;
                            'Backend Type'              = $null;
                            'Probe Name'                = $null;
                            'Probe Interval (sec)'      = $null;
                            'Probe Protocol'            = $null;
                            'Probe Port'                = $null;
                            'Probe Unhealthy threshold' = $null;
                            'Resource U'                = $ResUCount;
                            'Tag Name'                  = [string]$TagKey;
                            'Tag Value'                 = [string]$Tag.$TagKey
                        }
                        $tmp += $obj   
                        if ($ResUCount -eq 1) { $ResUCount = 0 }      
                    }
                }
                else {
                    $obj = @{
                        'Subscription'              = $sub1.name;
                        'Resource Group'            = $1.RESOURCEGROUP;
                        'Name'                      = $1.NAME;
                        'Location'                  = $1.LOCATION;
                        'SKU'                       = $1.sku.name;
                        'Frontend Name'             = $2.name;
                        'Frontend Target'           = $Fronttarget;
                        'Frontend Type'             = $FrontType;
                        'Frontend Subnet'           = $frontsub;
                        'Backend Pool Name'         = $null;
                        'Backend Target'            = $null;
                        'Backend Type'              = $null;
                        'Probe Name'                = $null;
                        'Probe Interval (sec)'      = $null;
                        'Probe Protocol'            = $null;
                        'Probe Port'                = $null;
                        'Probe Unhealthy threshold' = $null;
                        'Resource U'                = $ResUCount;
                        'Tag Name'                  = $null;
                        'Tag Value'                 = $null
                    }
                    $tmp += $obj 
                    if ($ResUCount -eq 1) { $ResUCount = 0 }  
                }     
            }
        }   
        elseif ($null -ne $data.frontendIPConfigurations -and $null -eq $data.backendAddressPools -and $null -ne $data.probes) {
            foreach ($2 in $data.frontendIPConfigurations) {
                $Fronttarget = ''    
                $Frontsub = ''
                if ($null -ne $2.properties.subnet.id) {
                    $Fronttarget = $2.properties.subnet.id.split('/')[8]
                    $Frontsub = $2.properties.subnet.id.split('/')[10]
                    $FrontType = 'VNET' 
                }
                elseif ($null -ne $2.properties.publicIPAddress.id) {
                    $Fronttarget = $2.properties.publicIPAddress.id.split('/')[8]
                    $Frontsub = ''
                    $FrontType = 'Public IP' 
                }        
                foreach ($3 in $data.probes) {
                    if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
                        foreach ($TagKey in $Tag.Keys) {
                            $obj = @{
                                'Subscription'              = $sub1.name;
                                'Resource Group'            = $1.RESOURCEGROUP;
                                'Name'                      = $1.NAME;
                                'Location'                  = $1.LOCATION;
                                'SKU'                       = $1.sku.name;
                                'Frontend Name'             = $2.name;
                                'Frontend Target'           = $Fronttarget;
                                'Frontend Type'             = $FrontType;
                                'Frontend Subnet'           = $frontsub;
                                'Backend Pool Name'         = $null;
                                'Backend Target'            = $null;
                                'Backend Type'              = $null;
                                'Probe Name'                = $3.name;
                                'Probe Interval (sec)'      = $3.properties.intervalInSeconds;
                                'Probe Protocol'            = $3.properties.protocol;
                                'Probe Port'                = $3.properties.port;
                                'Probe Unhealthy threshold' = $3.properties.numberOfProbes;
                                'Resource U'                = $ResUCount;
                                'Tag Name'                  = [string]$TagKey;
                                'Tag Value'                 = [string]$Tag.$TagKey
                            }
                            $tmp += $obj  
                            if ($ResUCount -eq 1) { $ResUCount = 0 }     
                        }
                    }
                    else {  
                        $obj = @{
                            'Subscription'              = $sub1.name;
                            'Resource Group'            = $1.RESOURCEGROUP;
                            'Name'                      = $1.NAME;
                            'Location'                  = $1.LOCATION;
                            'SKU'                       = $1.sku.name;
                            'Frontend Name'             = $2.name;
                            'Frontend Target'           = $Fronttarget;
                            'Frontend Type'             = $FrontType;
                            'Frontend Subnet'           = $frontsub;
                            'Backend Pool Name'         = $null;
                            'Backend Target'            = $null;
                            'Backend Type'              = $null;
                            'Probe Name'                = $3.name;
                            'Probe Interval (sec)'      = $3.properties.intervalInSeconds;
                            'Probe Protocol'            = $3.properties.protocol;
                            'Probe Port'                = $3.properties.port;
                            'Probe Unhealthy threshold' = $3.properties.numberOfProbes;
                            'Resource U'                = $ResUCount;
                            'Tag Name'                  = $null;
                            'Tag Value'                 = $null
                        }
                        $tmp += $obj
                        if ($ResUCount -eq 1) { $ResUCount = 0 } 
                    }      
                }
            }
        }   
        elseif ($null -eq $data.frontendIPConfigurations -and $null -ne $data.backendAddressPools -and $null -ne $data.probes) {
            foreach ($2 in $data.backendAddressPools) {
                $BackTarget = ''
                $BackType = ''
                if ($null -ne $3.properties.backendIPConfigurations.id) {
                    $BackTarget = $2.properties.backendIPConfigurations.id.split('/')[8]
                    $BackType = $2.properties.backendIPConfigurations.id.split('/')[7]
                }
                foreach ($3 in $data.probes) {
                    if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
                        foreach ($TagKey in $Tag.Keys) {
                            $obj = @{
                                'Subscription'              = $sub1.name;
                                'Resource Group'            = $1.RESOURCEGROUP;
                                'Name'                      = $1.NAME;
                                'Location'                  = $1.LOCATION;
                                'SKU'                       = $1.sku.name;
                                'Frontend Name'             = $null;
                                'Frontend Target'           = $null;
                                'Frontend Type'             = $null;
                                'Frontend Subnet'           = $null;
                                'Backend Pool Name'         = $2.name;
                                'Backend Target'            = $BackTarget;
                                'Backend Type'              = $BackType;
                                'Probe Name'                = $3.name;
                                'Probe Interval (sec)'      = $3.properties.intervalInSeconds;
                                'Probe Protocol'            = $3.properties.protocol;
                                'Probe Port'                = $3.properties.port;
                                'Probe Unhealthy threshold' = $3.properties.numberOfProbes;
                                'Resource U'                = $ResUCount;
                                'Tag Name'                  = [string]$TagKey;
                                'Tag Value'                 = [string]$Tag.$TagKey
                            }
                            $tmp += $obj   
                            if ($ResUCount -eq 1) { $ResUCount = 0 }     
                        }
                    }
                    else { 
                        $obj = @{
                            'Subscription'              = $sub1.name;
                            'Resource Group'            = $1.RESOURCEGROUP;
                            'Name'                      = $1.NAME;
                            'Location'                  = $1.LOCATION;
                            'SKU'                       = $1.sku.name;
                            'Frontend Name'             = $null;
                            'Frontend Target'           = $null;
                            'Frontend Type'             = $null;
                            'Frontend Subnet'           = $null;
                            'Backend Pool Name'         = $2.name;
                            'Backend Target'            = $BackTarget;
                            'Backend Type'              = $BackType;
                            'Probe Name'                = $3.name;
                            'Probe Interval (sec)'      = $3.properties.intervalInSeconds;
                            'Probe Protocol'            = $3.properties.protocol;
                            'Probe Port'                = $3.properties.port;
                            'Probe Unhealthy threshold' = $3.properties.numberOfProbes;
                            'Resource U'                = $ResUCount;
                            'Tag Name'                  = $null;
                            'Tag Value'                 = $null
                        }
                        $tmp += $obj   
                        if ($ResUCount -eq 1) { $ResUCount = 0 } 
                    }     
                }
            }            
        }    
        elseif ($null -eq $data.frontendIPConfigurations -and $null -eq $data.backendAddressPools -and $null -ne $data.probes) {
            foreach ($2 in $data.probes) {
                if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
                    foreach ($TagKey in $Tag.Keys) {
                        $obj = @{
                            'Subscription'              = $sub1.name;
                            'Resource Group'            = $1.RESOURCEGROUP;
                            'Name'                      = $1.NAME;
                            'Location'                  = $1.LOCATION;
                            'SKU'                       = $1.sku.name;
                            'Frontend Name'             = $null;
                            'Frontend Target'           = $null;
                            'Frontend Type'             = $null;
                            'Frontend Subnet'           = $null;
                            'Backend Pool Name'         = $null;
                            'Backend Target'            = $null;
                            'Backend Type'              = $null;
                            'Probe Name'                = $2.name;
                            'Probe Interval (sec)'      = $2.properties.intervalInSeconds;
                            'Probe Protocol'            = $2.properties.protocol;
                            'Probe Port'                = $2.properties.port;
                            'Probe Unhealthy threshold' = $2.properties.numberOfProbes;
                            'Resource U'                = $ResUCount;
                            'Tag Name'                  = [string]$TagKey;
                            'Tag Value'                 = [string]$Tag.$TagKey
                        }
                        $tmp += $obj
                        if ($ResUCount -eq 1) { $ResUCount = 0 } 
                    }
                }
                else { 
                    $obj = @{
                        'Subscription'              = $sub1.name;
                        'Resource Group'            = $1.RESOURCEGROUP;
                        'Name'                      = $1.NAME;
                        'Location'                  = $1.LOCATION;
                        'SKU'                       = $1.sku.name;
                        'Frontend Name'             = $null;
                        'Frontend Target'           = $null;
                        'Frontend Type'             = $null;
                        'Frontend Subnet'           = $null;
                        'Backend Pool Name'         = $null;
                        'Backend Target'            = $null;
                        'Backend Type'              = $null;
                        'Probe Name'                = $2.name;
                        'Probe Interval (sec)'      = $2.properties.intervalInSeconds;
                        'Probe Protocol'            = $2.properties.protocol;
                        'Probe Port'                = $2.properties.port;
                        'Probe Unhealthy threshold' = $2.properties.numberOfProbes;
                        'Resource U'                = $ResUCount;
                        'Tag Name'                  = $null;
                        'Tag Value'                 = $null
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
    if ($SmaResources.LoadBalancer) {
        $txtLB = New-ConditionalText Basic -Range E:E
                        

        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat 0

        $ExcelLB = $SmaResources.LoadBalancer     

        if ($InTag -eq $True) {
            $ExcelLB | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'SKU',
            'Frontend Name',
            'Frontend Target',
            'Frontend Type',
            'Frontend Subnet',
            'Backend Pool Name',
            'Backend Target',
            'Backend Type',
            'Probe Name',
            'Probe Interval (sec)',
            'Probe Protocol',
            'Probe Port',
            'Probe Unhealthy threshold',
            'Tag Name',
            'Tag Value'  | 
            Export-Excel -Path $File -WorksheetName 'Load Balancers' -AutoSize -TableName 'LoadBalancers' -TableStyle $tableStyle -ConditionalText $txtLB -Style $Style
        }
        else {
            $ExcelLB | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'SKU',
            'Frontend Name',
            'Frontend Target',
            'Frontend Type',
            'Frontend Subnet',
            'Backend Pool Name',
            'Backend Target',
            'Backend Type',
            'Probe Name',
            'Probe Interval (sec)',
            'Probe Protocol',
            'Probe Port',
            'Probe Unhealthy threshold' | 
            Export-Excel -Path $File -WorksheetName 'Load Balancers' -AutoSize -TableName 'LoadBalancers' -TableStyle $tableStyle -ConditionalText $txtLB -Style $Style
        }

    }
}