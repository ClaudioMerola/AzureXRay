param($SCPath, $Sub, $Intag, $Resources, $Task , $File, $SmaResources, $TableStyle) 
If ($Task -eq 'Processing') {

    $PublicDNS = $Resources | Where-Object { $_.TYPE -eq 'microsoft.network/dnszones' }

    $tmp = @()
    $Subs = $Sub

    foreach ($1 in $PublicDNS) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
            foreach ($TagKey in $Tag.Keys) {     
                $obj = @{
                    'Subscription'              = $sub1.name;
                    'Resource Group'            = $1.RESOURCEGROUP;
                    'Name'                      = $1.NAME;
                    'Location'                  = $1.LOCATION;
                    'Zone Type'                 = $data.zoneType;
                    'Number of Record Sets'     = $data.numberOfRecordSets;
                    'Max Number of Record Sets' = $data.maxNumberofRecordSets;
                    'Name Servers'              = [string]$data.nameServers;
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
                'Zone Type'                 = $data.zoneType;
                'Number of Record Sets'     = $data.numberOfRecordSets;
                'Max Number of Record Sets' = $data.maxNumberofRecordSets;
                'Name Servers'              = [string]$data.nameServers;
                'Resource U'                = $ResUCount;
                'Tag Name'                  = $null;
                'Tag Value'                 = $null
            }
            $tmp += $obj
            if ($ResUCount -eq 1) { $ResUCount = 0 } 
        }
    }
    $tmp
}
Else {
    if ($SmaResources.PublicDNS) {
        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat 0

        $ExcelPubDNS = $AzNetwork.PublicDNS

        if ($InTag -eq $True) {
            $ExcelPubDNS | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'Zone Type',
            'Number of Record Sets',
            'Max Number of Record Sets',
            'Name Servers',
            'Tag Name',
            'Tag Value' | 
            Export-Excel -Path $File -WorksheetName 'Public DNS' -AutoSize -TableName 'AzurePubDNSZones' -TableStyle $tableStyle -Style $Style
        }
        else {
            $ExcelPubDNS | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'Zone Type',
            'Number of Record Sets',
            'Max Number of Record Sets',
            'Name Servers' | 
            Export-Excel -Path $File -WorksheetName 'Public DNS' -AutoSize -TableName 'AzurePubDNSZones' -TableStyle $tableStyle -Style $Style
        }
    }
}