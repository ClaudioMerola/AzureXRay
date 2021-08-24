<######## Default Parameters. Don't modify this ########>

param($SCPath, $Sub, $Intag, $Resources, $Task ,$File, $SmaResources, $TableStyle)
 
If ($Task -eq 'Processing')
{
 
    <######### Insert the resource extraction here ########>

        $wrkspace = $Resources | Where-Object {$_.TYPE -eq 'microsoft.operationalinsights/workspaces'}

    <######### Insert the resource Process here ########>

    $tmp = @()
    $Subs = $Sub

    foreach ($1 in $wrkspace) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
            foreach ($TagKey in $Tag.Keys) {
                $obj = @{
                    'Subscription'     = $sub1.name;
                    'Resource Group'   = $1.RESOURCEGROUP;
                    'Name'             = $1.NAME;
                    'Location'         = $1.LOCATION;
                    'SKU'              = $data.sku.name;
                    'Retention Days'   = $data.retentionInDays;
                    'Daily Quota (GB)' = [decimal]$data.workspaceCapping.dailyQuotaGb;
                    'Resource U'       = $ResUCount;
                    'Tag Name'         = [string]$TagKey;
                    'Tag Value'        = [string]$Tag.$TagKey
                }
                $tmp += $obj
                if ($ResUCount -eq 1) { $ResUCount = 0 } 
            }
        }
        else {
            $obj = @{
                'Subscription'     = $sub1.name;
                'Resource Group'   = $1.RESOURCEGROUP;
                'Name'             = $1.NAME;
                'Location'         = $1.LOCATION;
                'SKU'              = $data.sku.name;
                'Retention Days'   = $data.retentionInDays;
                'Daily Quota (GB)' = [decimal]$data.workspaceCapping.dailyQuotaGb;
                'Resource U'       = $ResUCount;
                'Tag Name'         = $null;
                'Tag Value'        = $null
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

    if($SmaResources.WrkSpace)
    {

        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat '0.0'
            
        $ExcelWrkSpace = $SmaResources.WrkSpace

        if ($InTag -eq $true) {
            $ExcelWrkSpace | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'SKU',
            'Retention Days',
            'Daily Quota (GB)',
            'Tag Name',
            'Tag Value' | 
            Export-Excel -Path $File -WorksheetName 'Workspaces' -AutoSize -TableName 'AzureWorkspace' -TableStyle $tableStyle -Style $Style
        }
        else {
            $ExcelWrkSpace | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'SKU',
            'Retention Days',
            'Daily Quota (GB)' | 
            Export-Excel -Path $File -WorksheetName 'Workspaces' -AutoSize -TableName 'AzureWorkspace' -TableStyle $tableStyle -Style $Style
        }

        <######## Insert Column comments and documentations here following this model #########>


        #$excel = Open-ExcelPackage -Path $File -KillExcel


        #Close-ExcelPackage $excel 

    }
}