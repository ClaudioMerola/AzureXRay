<######## Default Parameters. Don't modify this ########>

param($SCPath, $Sub, $Intag, $Resources, $Task ,$File, $SmaResources, $TableStyle)
 
If ($Task -eq 'Processing')
{
 
    <######### Insert the resource extraction here ########>

        $evthub = $Resources | Where-Object {$_.TYPE -eq 'microsoft.eventhub/namespaces'}

    <######### Insert the resource Process here ########>

    $tmp = @()
    $Subs = $Sub

    foreach ($1 in $evthub) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $sku = $1.SKU
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
            foreach ($TagKey in $Tag.Keys) { 
                $obj = @{
                    'Subscription'         = $sub1.name;
                    'Resource Group'       = $1.RESOURCEGROUP;
                    'Name'                 = $1.NAME;
                    'Location'             = $1.LOCATION;
                    'SKU'                  = $sku.name;
                    'Status'               = $data.status;
                    'Geo-Replication'      = $data.zoneRedundant;
                    'Throughput Units'     = $1.sku.capacity;
                    'Auto-Inflate'         = $data.isAutoInflateEnabled;
                    'Max Throughput Units' = $data.maximumThroughputUnits;
                    'Kafka Enabled'        = $data.kafkaEnabled;
                    'Endpoint'             = $data.serviceBusEndpoint;
                    'Resource U'           = $ResUCount;
                    'Tag Name'             = [string]$TagKey;
                    'Tag Value'            = [string]$Tag.$TagKey
                }
                $tmp += $obj
                if ($ResUCount -eq 1) { $ResUCount = 0 } 
            }
        }
        else { 
            $obj = @{
                'Subscription'         = $sub1.name;
                'Resource Group'       = $1.RESOURCEGROUP;
                'Name'                 = $1.NAME;
                'Location'             = $1.LOCATION;
                'SKU'                  = $sku.name;
                'Status'               = $data.status;
                'Geo-Replication'      = $data.zoneRedundant;
                'Throughput Units'     = $1.sku.capacity;
                'Auto-Inflate'         = $data.isAutoInflateEnabled;
                'Max Throughput Units' = $data.maximumThroughputUnits;
                'Kafka Enabled'        = $data.kafkaEnabled;
                'Endpoint'             = $data.serviceBusEndpoint;
                'Resource U'           = $ResUCount;
                'Tag Name'             = $null;
                'Tag Value'            = $null
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

    if($SmaResources.EvtHub)
    {

        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat '0'

        $txtEvt = $(New-ConditionalText false -Range I:I
            New-ConditionalText falso -Range I:I)

        $ExcelEvtHub = $SmaResources.EvtHub

        if ($InTag -eq $true) {
            $ExcelEvtHub | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'SKU',
            'Status',
            'Geo-Rep',
            'Throughput Units',
            'Auto-Inflate',
            'Max Throughput Units',
            'Kafka Enabled',
            'Endpoint',
            'Tag Name',
            'Tag Value' | 
            Export-Excel -Path $File -WorksheetName 'Event Hubs' -AutoSize -TableName 'AzureEventHubs' -TableStyle $tableStyle -ConditionalText $txtEvt -Style $Style
        }
        else {
            $ExcelEvtHub | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'SKU',
            'Status',
            'Geo-Rep',
            'Throughput Units',
            'Auto-Inflate',
            'Max Throughput Units',
            'Kafka Enabled',
            'Endpoint' | 
            Export-Excel -Path $File -WorksheetName 'Event Hubs' -AutoSize -TableName 'AzureEventHubs' -TableStyle $tableStyle -ConditionalText $txtEvt -Style $Style
        }

        <######## Insert Column comments and documentations here following this model #########>


        #$excel = Open-ExcelPackage -Path $File -KillExcel


        #Close-ExcelPackage $excel 

    }
}