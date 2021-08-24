<######## Default Parameters. Don't modify this ########>

param($SCPath, $Sub, $Intag, $Resources, $Task ,$File, $SmaResources, $TableStyle)
 
If ($Task -eq 'Processing')
{
 
    <######### Insert the resource extraction here ########>

        $webfarm = $Resources | Where-Object {$_.TYPE -eq 'microsoft.web/serverfarms'}

    <######### Insert the resource Process here ########>

    $tmp = @()
    $Subs = $Sub

    foreach ($1 in $webfarm) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $sku = $1.SKU
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
            foreach ($TagKey in $Tag.Keys) {
                $obj = @{
                    'Subscription'        = $sub1.name;
                    'Resource Group'      = $1.RESOURCEGROUP;
                    'Name'                = $1.NAME;
                    'Location'            = $1.LOCATION;
                    'SKU'                 = $sku.name;
                    'SKU Family'          = $sku.family;
                    'Tier'                = $sku.tier;
                    'Capacity'            = $sku.capacity;
                    'Workers'             = $data.currentNumberOfWorkers;
                    'Compute Mode'        = $data.computeMode;
                    'Max Elastic Workers' = $data.maximumElasticWorkerCount;
                    'Max Workers'         = $data.maximumNumberOfWorkers;
                    'Worker Kind'         = $data.kind;
                    'Number Of Sites'     = $data.numberOfSites;
                    'Plan Name'           = $data.planName;
                    'Resource U'          = $ResUCount;
                    'Tag Name'            = [string]$TagKey;
                    'Tag Value'           = [string]$Tag.$TagKey
                }
                $tmp += $obj
                if ($ResUCount -eq 1) { $ResUCount = 0 } 
            }
        }
        else {
            $obj = @{
                'Subscription'        = $sub1.name;
                'Resource Group'      = $1.RESOURCEGROUP;
                'Name'                = $1.NAME;
                'Location'            = $1.LOCATION;
                'SKU'                 = $sku.name;
                'SKU Family'          = $sku.family;
                'Tier'                = $sku.tier;
                'Capacity'            = $sku.capacity;
                'Workers'             = $data.currentNumberOfWorkers;
                'Compute Mode'        = $data.computeMode;
                'Max Elastic Workers' = $data.maximumElasticWorkerCount;
                'Max Workers'         = $data.maximumNumberOfWorkers;
                'Worker Kind'         = $data.kind;
                'Number Of Sites'     = $data.numberOfSites;
                'Plan Name'           = $data.planName;
                'Resource U'          = $ResUCount;
                'Tag Name'            = $null;
                'Tag Value'           = $null
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

    if($SmaResources.WEBSERVERFARM)
    {

        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat '0'

        $ExcelWebFarm = $SmaResources.WEBSERVERFARM

        if ($InTag -eq $true) {
            $ExcelWebFarm | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'SKU',
            'SKU Family',
            'Tier',
            'Capacity',
            'Workers',
            'Compute Mode',
            'Max Elastic Workers',
            'Max Workers',
            'Worker Kind',
            'Number Of Sites',
            'Plan Name',
            'Tag Name',
            'Tag Value' | 
            Export-Excel -Path $File -WorksheetName 'Web Servers' -AutoSize -TableName 'AzureWebServers' -TableStyle $tableStyle -Style $Style
        }
        else {
            $ExcelWebFarm | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'SKU',
            'SKU Family',
            'Tier',
            'Capacity',
            'Workers',
            'Compute Mode',
            'Max Elastic Workers',
            'Max Workers',
            'Worker Kind',
            'Number Of Sites',
            'Plan Name' | 
            Export-Excel -Path $File -WorksheetName 'Web Servers' -AutoSize -TableName 'AzureWebServers' -TableStyle $tableStyle -Style $Style
        }

        <######## Insert Column comments and documentations here following this model #########>


        #$excel = Open-ExcelPackage -Path $File -KillExcel


        #Close-ExcelPackage $excel 

    }
}