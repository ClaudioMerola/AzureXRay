<######## Default Parameters. Don't modify this ########>

param($SCPath, $Sub, $Intag, $Resources, $Task ,$File, $SmaResources, $TableStyle)
 
If ($Task -eq 'Processing')
{
 
    <######### Insert the resource extraction here ########>

        $RECOVAULT = $Resources | Where-Object {$_.TYPE -eq 'microsoft.recoveryservices/vaults'}

    <######### Insert the resource Process here ########>

    $tmp = @()

    $Subs = $Sub

    foreach ($1 in $RECOVAULT) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
            foreach ($TagKey in $Tag.Keys) {
                $obj = @{
                    'Subscription'                             = $sub1.name;
                    'Resource Group'                           = $1.RESOURCEGROUP;
                    'Name'                                     = $1.NAME;
                    'Location'                                 = $1.LOCATION;
                    'SKU Name'                                 = $1.sku.name;
                    'SKU Tier'                                 = $1.sku.tier;
                    'Private Endpoint State for Backup'        = $data.privateEndpointStateForBackup;
                    'Private Endpoint State for Site Recovery' = $data.privateEndpointStateForSiteRecovery;
                    'Resource U'                               = $ResUCount;
                    'Tag Name'                                 = [string]$TagKey;
                    'Tag Value'                                = [string]$Tag.$TagKey
                }
                $tmp += $obj
                if ($ResUCount -eq 1) { $ResUCount = 0 } 
            }
        }
        else {
            $obj = @{
                'Subscription'                             = $sub1.name;
                'Resource Group'                           = $1.RESOURCEGROUP;
                'Name'                                     = $1.NAME;
                'Location'                                 = $1.LOCATION;
                'SKU Name'                                 = $1.sku.name;
                'SKU Tier'                                 = $1.sku.tier;
                'Private Endpoint State for Backup'        = $data.privateEndpointStateForBackup;
                'Private Endpoint State for Site Recovery' = $data.privateEndpointStateForSiteRecovery;
                'Resource U'                               = $ResUCount;
                'Tag Name'                                 = $null;
                'Tag Value'                                = $null
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

    if($SmaResources.RecoveryVault)
    {

        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat '0'

        $ExcelRecVault = $SmaResources.RecoveryVault

        if ($InTag -eq $true) {
            $ExcelRecVault | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'SKU Name',
            'SKU Tier',
            'Private Endpoint State for Backup',
            'Private Endpoint State for Site Recovery',
            'Tag Name',
            'Tag Value' | 
            Export-Excel -Path $File -WorksheetName 'Recovery Vaults' -AutoSize -TableName 'AzureRecVault' -TableStyle $tableStyle -Style $Style
        }
        else {
            $ExcelRecVault | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'SKU Name',
            'SKU Tier',
            'Private Endpoint State for Backup',
            'Private Endpoint State for Site Recovery' | 
            Export-Excel -Path $File -WorksheetName 'Recovery Vaults' -AutoSize -TableName 'AzureRecVault' -TableStyle $tableStyle -Style $Style
        }

        <######## Insert Column comments and documentations here following this model #########>


        #$excel = Open-ExcelPackage -Path $File -KillExcel


        #Close-ExcelPackage $excel 

    }
}