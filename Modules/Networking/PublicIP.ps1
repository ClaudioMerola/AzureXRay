param($SCPath, $Sub, $Intag, $Resources, $Task , $File, $SmaResources, $TableStyle) 
If ($Task -eq 'Processing') {

    $PublicIP = $Resources | Where-Object { $_.TYPE -eq 'microsoft.network/publicipaddresses' }

    $tmp = @()
    $Subs = $Sub

    foreach ($1 in $PublicIP) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $Tag = @{}
        if (!($data.ipConfiguration.id)) { $Use = 'Underutilized' } else { $Use = 'Utilized' }
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        if ($null -ne $data.ipConfiguration.id -and ![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
            foreach ($TagKey in $Tag.Keys) { 
                $obj = @{
                    'Subscription'             = $sub1.name;
                    'Resource Group'           = $1.RESOURCEGROUP;
                    'Name'                     = $1.NAME;
                    'SKU'                      = $1.SKU.Name;
                    'Location'                 = $1.LOCATION;
                    'Type'                     = $data.publicIPAllocationMethod;
                    'Version'                  = $data.publicIPAddressVersion;
                    'IP Address'               = $data.ipAddress;
                    'Use'                      = $Use;
                    'Associated Resource'      = $data.ipConfiguration.id.split('/')[8];
                    'Associated Resource Type' = $data.ipConfiguration.id.split('/')[7];
                    'Resource U'               = $ResUCount;
                    'Tag Name'                 = [string]$TagKey;
                    'Tag Value'                = [string]$Tag.$TagKey
                }
                $tmp += $obj
                if ($ResUCount -eq 1) { $ResUCount = 0 } 
            }
        }
        elseif ($null -ne $data.ipConfiguration.id -and $InTag -ne $true) { 
            $obj = @{
                'Subscription'             = $sub1.name;
                'Resource Group'           = $1.RESOURCEGROUP;
                'Name'                     = $1.NAME;
                'SKU'                      = $1.SKU.Name;
                'Location'                 = $1.LOCATION;
                'Type'                     = $data.publicIPAllocationMethod;
                'Version'                  = $data.publicIPAddressVersion;
                'IP Address'               = $data.ipAddress;
                'Use'                      = $Use;
                'Associated Resource'      = $data.ipConfiguration.id.split('/')[8];
                'Associated Resource Type' = $data.ipConfiguration.id.split('/')[7];
                'Resource U'               = $ResUCount;
                'Tag Name'                 = $null;
                'Tag Value'                = $null
            }
            $tmp += $obj
            if ($ResUCount -eq 1) { $ResUCount = 0 } 
        }
        elseif ($null -eq $data.ipConfiguration.id -and ![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
            foreach ($TagKey in $Tag.Keys) {  
                $obj = @{
                    'Subscription'             = $sub1.name;
                    'Resource Group'           = $1.RESOURCEGROUP;
                    'Name'                     = $1.NAME;
                    'SKU'                      = $1.SKU.Name;
                    'Location'                 = $1.LOCATION;
                    'Type'                     = $data.publicIPAllocationMethod;
                    'Version'                  = $data.publicIPAddressVersion;
                    'IP Address'               = $data.ipAddress;
                    'Use'                      = $Use;
                    'Associated Resource'      = $null;
                    'Associated Resource Type' = $null;
                    'Resource U'               = $ResUCount;
                    'Tag Name'                 = [string]$TagKey;
                    'Tag Value'                = [string]$Tag.$TagKey
                }
                $tmp += $obj
                if ($ResUCount -eq 1) { $ResUCount = 0 } 
            }
        }
        elseif ($null -ne $data.ipConfiguration.id -and [string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
            $obj = @{
                'Subscription'             = $sub1.name;
                'Resource Group'           = $1.RESOURCEGROUP;
                'Name'                     = $1.NAME;
                'SKU'                      = $1.SKU.Name;
                'Location'                 = $1.LOCATION;
                'Type'                     = $data.publicIPAllocationMethod;
                'Version'                  = $data.publicIPAddressVersion;
                'IP Address'               = $data.ipAddress;
                'Use'                      = $Use;
                'Associated Resource'      = $data.ipConfiguration.id.split('/')[8];
                'Associated Resource Type' = $data.ipConfiguration.id.split('/')[7];
                'Resource U'               = $ResUCount;
                'Tag Name'                 = $null;
                'Tag Value'                = $null
            }
            $tmp += $obj
            if ($ResUCount -eq 1) { $ResUCount = 0 } 
        }
        elseif ($null -eq $data.ipConfiguration.id -and $InTag -ne $true) {  
            $obj = @{
                'Subscription'             = $sub1.name;
                'Resource Group'           = $1.RESOURCEGROUP;
                'Name'                     = $1.NAME;
                'SKU'                      = $1.SKU.Name;
                'Location'                 = $1.LOCATION;
                'Type'                     = $data.publicIPAllocationMethod;
                'Version'                  = $data.publicIPAddressVersion;
                'IP Address'               = $data.ipAddress;
                'Use'                      = $Use;
                'Associated Resource'      = $null;
                'Associated Resource Type' = $null;
                'Resource U'               = $ResUCount;
                'Tag Name'                 = $null;
                'Tag Value'                = $null
            }
            $tmp += $obj
            if ($ResUCount -eq 1) { $ResUCount = 0 } 
        }
    }
    $tmp
}
Else {
    if ($SmaResources.PublicIP) {
        $condtxtpip = New-ConditionalText Underutilized -Range I:I
        Write-Debug ('Generating Public IP sheet for: ' + $pubip.count + ' Public IPs.')

        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat 0

        $ExcelPIP = $SmaResources.PIP     

        if ($InTag -eq $True) {
            $ExcelPIP | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'SKU',
            'Location',
            'Type',
            'Version',
            'IP Address',
            'Use',
            'Associated Resource',
            'Associated Resource Type',
            'Tag Name',
            'Tag Value' | 
            Export-Excel -Path $File -WorksheetName 'Public IPs' -AutoSize -TableName 'AzurePubIPs' -TableStyle $tableStyle -Style $Style -ConditionalText $condtxtpip
        }
        else {
            $ExcelPIP | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'SKU',
            'Location',
            'Type',
            'Version',
            'IP Address',
            'Use',
            'Associated Resource',
            'Associated Resource Type' | 
            Export-Excel -Path $File -WorksheetName 'Public IPs' -AutoSize -TableName 'AzurePubIPs' -TableStyle $tableStyle -Style $Style -ConditionalText $condtxtpip
        }
    }
}
