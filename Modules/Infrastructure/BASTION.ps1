<######## Default Parameters. Don't modify this ########>

param($SCPath, $Sub, $Intag, $Resources, $Task ,$File, $SmaResources, $TableStyle)
 
If ($Task -eq 'Processing')
{
 
    <######### Insert the resource extraction here ########>

        $BASTION = $Resources | Where-Object {$_.TYPE -eq 'microsoft.network/bastionhosts'}

    <######### Insert the resource Process here ########>

    $tmp = @()
    $Subs = $Sub

    foreach ($1 in $BASTION) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $BastVNET = $data.ipConfigurations.properties.subnet.id.split("/")[8]
        $BastPIP = $data.ipConfigurations.properties.publicIPAddress.id.split("/")[8]
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
            foreach ($TagKey in $Tag.Keys) {
                $obj = @{
                    'Subscription'    = $sub1.name;
                    'Resource Group'  = $1.RESOURCEGROUP;
                    'Name'            = $1.NAME;
                    'Location'        = $1.LOCATION;
                    'SKU'             = $1.sku.name;
                    'DNS Name'        = $data.dnsName;
                    'Virtual Network' = $BastVNET;
                    'Public IP'       = $BastPIP;
                    'Scale Units'     = $data.scaleUnits;
                    'Tag Name'        = [string]$TagKey;
                    'Tag Value'       = [string]$Tag.$TagKey
                }
                $tmp += $obj
                if ($ResUCount -eq 1) { $ResUCount = 0 } 
            }
        }
        else {
            $obj = @{
                'Subscription'    = $sub1.name;
                'Resource Group'  = $1.RESOURCEGROUP;
                'Name'            = $1.NAME;
                'Location'        = $1.LOCATION;
                'SKU'             = $1.sku.name;
                'DNS Name'        = $data.dnsName;
                'Virtual Network' = $BastVNET;
                'Public IP'       = $BastPIP;
                'Scale Units'     = $data.scaleUnits;
                'Tag Name'        = $null;
                'Tag Value'       = $null
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

    if($SmaResources.BASTION)
    {

        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat '0'

        $ExcelBASTION = $SmaResources.BASTION

        if ($InTag -eq $true) {
            $ExcelBASTION | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'SKU',
            'DNS Name',
            'Virtual Network',
            'Public IP',
            'Scale Units',
            'Tag Name',
            'Tag Value' | 
            Export-Excel -Path $File -WorksheetName 'Bastion Hosts' -AutoSize -TableName 'AzureBastion' -TableStyle $tableStyle -Style $Style
        }
        else {
            $ExcelBASTION | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'SKU',
            'DNS Name',
            'Virtual Network',
            'Public IP',
            'Scale Units' | 
            Export-Excel -Path $File -WorksheetName 'Bastion Hosts' -AutoSize -TableName 'AzureBastion' -TableStyle $tableStyle -Style $Style
        }

        <######## Insert Column comments and documentations here following this model #########>


        #$excel = Open-ExcelPackage -Path $File -KillExcel


        #Close-ExcelPackage $excel 

    }
}