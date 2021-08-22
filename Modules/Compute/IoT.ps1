<######## Default Parameters. Don't modify this ########>

param($SCPath, $Sub, $Intag, $Resources, $Task ,$File, $SmaResources, $TableStyle)
 
If ($Task -eq 'Processing')
{
 
    <######### Insert the resource extraction here ########>

    $IoT = @()

    ForEach ($Resource in $Resources) {          
            if ($Resource.TYPE -eq 'microsoft.devices/iothubs' ) { $IoT += $Resource }
    }

    <######### Insert the resource Process here ########>

    $tmp = @()

    $Subs = $Sub
    
    foreach ($1 in $IoT) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
            foreach ($TagKey in $Tag.Keys) {
                $obj = @{
                    'Subscription'                     = $sub1.name;
                    'Resource Group'                   = $1.RESOURCEGROUP;
                    'Name'                             = $1.NAME;
                    'HostName'                         = $data.hostname;
                    'State'                            = $data.state;
                    'SKU'                              = $1.sku.name;
                    'SKU Tier'                         = $1.sku.tier;
                    'SKU Capacity'                     = $1.sku.capacity;
                    'Features'                         = $data.features;
                    'Enable File Upload Notifications' = $data.enableFileUploadNotifications;
                    'Default TTL As ISO8601'           = $data.cloudToDevice.defaultTtlAsIso8601;
                    'Max Delivery Count'               = $data.cloudToDevice.maxDeliveryCount;
                    'EventHubs Endpoint'               = $data.eventHubEndpoints.events.endpoint;
                    'EventHubs Partition Count'        = $data.eventHubEndpoints.events.partitionCount;
                    'EventHubs Path'                   = $data.eventHubEndpoints.events.path;
                    'EventHubs Retention Days'         = $data.eventHubEndpoints.events.retentionTimeInDays;
                    'Locations'                        = [string]$data.locations.location;
                    'Resource U'                       = $ResUCount;
                    'Tag Name'                         = [string]$TagKey;
                    'Tag Value'                        = [string]$Tag.$TagKey
                }
                $tmp += $obj
                if ($ResUCount -eq 1) { $ResUCount = 0 } 
            }
        }
        else {
            $obj = @{
                'Subscription'                     = $sub1.name;
                'Resource Group'                   = $1.RESOURCEGROUP;
                'Name'                             = $1.NAME;
                'HostName'                         = $data.hostname;
                'State'                            = $data.state;
                'SKU'                              = $1.sku.name;
                'SKU Tier'                         = $1.sku.tier;
                'SKU Capacity'                     = $1.sku.capacity;
                'Features'                         = $data.features;
                'Enable File Upload Notifications' = $data.enableFileUploadNotifications;
                'Default TTL As ISO8601'           = $data.cloudToDevice.defaultTtlAsIso8601;
                'Max Delivery Count'               = $data.cloudToDevice.maxDeliveryCount;
                'EventHubs Endpoint'               = $data.eventHubEndpoints.events.endpoint;
                'EventHubs Partition Count'        = $data.eventHubEndpoints.events.partitionCount;
                'EventHubs Path'                   = $data.eventHubEndpoints.events.path;
                'EventHubs Retention Days'         = $data.eventHubEndpoints.events.retentionTimeInDays;
                'Locations'                        = [string]$data.locations.location;
                'Resource U'                       = $ResUCount;
                'Tag Name'                         = $null;
                'Tag Value'                        = $null
            }
            $tmp += $obj
            if ($ResUCount -eq 1) { $ResUCount = 0 } 
        }
    }
    $tmp

}

<######## Resource Excel Reporting Begins Here ########>

Else
{
    <######## $SmaResources.(RESOURCE FILE NAME) ##########>

    if($SmaResources.IoT)
    {

        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat '0'

        $ExcelIot = $SmaResources.IoT

        if ($InTag -eq $true) {
            $ExcelIot | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'HostName',
            'State',
            'SKU',
            'SKU Tier',
            'SKU Capacity',
            'Features',
            'Enable File Upload Notifications',
            'Default TTL As ISO8601',
            'Max Delivery Count',
            'EventHubs Endpoint',
            'EventHubs Partition Count',
            'EventHubs Path',
            'EventHubs Retention Days',
            'Locations',
            'Tag Name',
            'Tag Value' | 
            Export-Excel -Path $File -WorksheetName 'IoT Hubs' -AutoSize -TableName 'AzureIOT' -TableStyle $tableStyle -Style $Style
        }
        else {
            $ExcelIot | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'HostName',
            'State',
            'SKU',
            'SKU Tier',
            'SKU Capacity',
            'Features',
            'Enable File Upload Notifications',
            'Default TTL As ISO8601',
            'Max Delivery Count',
            'EventHubs Endpoint',
            'EventHubs Partition Count',
            'EventHubs Path',
            'EventHubs Retention Days',
            'Locations' | 
            Export-Excel -Path $File -WorksheetName 'IoT Hubs' -AutoSize -TableName 'AzureIOT' -TableStyle $tableStyle -Style $Style
        }

        <######## Insert Column comments and documentations here following this model #########>


        #$excel = Open-ExcelPackage -Path $File -KillExcel


        #Close-ExcelPackage $excel 

    }
}