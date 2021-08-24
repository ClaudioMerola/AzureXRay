param($SCPath, $Sub, $Intag, $Resources, $Task , $File, $SmaResources, $TableStyle) 
If ($Task -eq 'Processing') {

    $PrivateDNS = $Resources | Where-Object { $_.TYPE -eq 'microsoft.network/privatednszones' }  

    $tmp = @()
    $Subs = $Sub

    foreach ($1 in $PrivateDNS) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
            foreach ($TagKey in $Tag.Keys) {     
                $obj = @{
                    'Subscription'                    = $sub1.name;
                    'Resource Group'                  = $1.RESOURCEGROUP;
                    'Name'                            = $1.NAME;
                    'Location'                        = $1.LOCATION;
                    'Number of Records'               = $data.numberOfRecordSets;
                    'Virtual Network Links'           = $data.numberOfVirtualNetworkLinks;
                    'Network Links with Registration' = $data.numberOfVirtualNetworkLinksWithRegistration;
                    'Tag Name'                        = [string]$TagKey;
                    'Tag Value'                       = [string]$Tag.$TagKey
                }
                $tmp += $obj
                if ($ResUCount -eq 1) { $ResUCount = 0 } 
            }
        }
        else {    
            $obj = @{
                'Subscription'                    = $sub1.name;
                'Resource Group'                  = $1.RESOURCEGROUP;
                'Name'                            = $1.NAME;
                'Location'                        = $1.LOCATION;
                'Number of Records'               = $data.numberOfRecordSets;
                'Virtual Network Links'           = $data.numberOfVirtualNetworkLinks;
                'Network Links with Registration' = $data.numberOfVirtualNetworkLinksWithRegistration;
                'Tag Name'                        = $null;
                'Tag Value'                       = $null
            }
            $tmp += $obj
            if ($ResUCount -eq 1) { $ResUCount = 0 } 
        }
    }
    $tmp
}
Else {
    if ($SmaResources.PrivateDNS) {
        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat 0

        $ExcelPrivDNS = $SmaResources.PrivateDNS

        if ($InTag -eq $True) {
            $ExcelPrivDNS | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'Number of Records',
            'Virtual Network Links',
            'Network Links with Registration',
            'Tag Name',
            'Tag Value' | 
            Export-Excel -Path $File -WorksheetName 'Private DNS' -AutoSize -TableName 'AzurePrivDNSZones' -TableStyle $tableStyle -Style $Style
        }
        else {
            $ExcelPrivDNS | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'Number of Records',
            'Virtual Network Links',
            'Network Links with Registration' | 
            Export-Excel -Path $File -WorksheetName 'Private DNS' -AutoSize -TableName 'AzurePrivDNSZones' -TableStyle $tableStyle -Style $Style
        }
    }   
}