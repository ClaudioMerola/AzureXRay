<######## Default Parameters. Don't modify this ########>

param($SCPath, $Sub, $Intag, $Resources, $Task , $File, $SmaResources, $TableStyle)

If ($Task -eq 'Processing') {

    <######### Insert the resource extraction here ########>

    $SQLVM = $Resources | Where-Object { $_.TYPE -eq 'microsoft.sqlvirtualmachine/sqlvirtualmachines' }

    $tmp = @()
    $Subs = $Sub

    foreach ($1 in $SQLVM) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
            foreach ($TagKey in $Tag.Keys) {
                $obj = @{
                    'Subscription'            = $sub1.name;
                    'ResourceGroup'           = $1.RESOURCEGROUP;
                    'Name'                    = $1.NAME;
                    'Location'                = $1.LOCATION;
                    'Zone'                    = $1.ZONES;
                    'SQL Server License Type' = $data.sqlServerLicenseType;
                    'SQL Image'               = $data.sqlImageOffer;
                    'SQL Management'          = $data.sqlManagement;
                    'SQL Image Sku'           = $data.sqlImageSku;
                    'Resource U'              = $ResUCount;
                    'Tag Name'                = [string]$TagKey;
                    'Tag Value'               = [string]$Tag.$TagKey
                }
                $tmp += $obj
                if ($ResUCount -eq 1) { $ResUCount = 0 } 
            }
        }
        else {
            $obj = @{
                'Subscription'            = $sub1.name;
                'ResourceGroup'           = $1.RESOURCEGROUP;
                'Name'                    = $1.NAME;
                'Location'                = $1.LOCATION;
                'Zone'                    = $1.ZONES;
                'SQL Server License Type' = $data.sqlServerLicenseType;
                'SQL Image'               = $data.sqlImageOffer;
                'SQL Management'          = $data.sqlManagement;
                'SQL Image Sku'           = $data.sqlImageSku;
                'Resource U'              = $ResUCount;
                'Tag Name'                = $null;
                'Tag Value'               = $null
            }
            $tmp += $obj
            if ($ResUCount -eq 1) { $ResUCount = 0 } 
        }
    }
    $tmp
}
<######## Resource Excel Reporting Begins Here ########>

Else {
    <######## $SmaResources.(RESOURCE FILE NAME) ##########>

    if ($SmaResources.SQLVM) {
        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat 0
        $ExcelSQLVM = $SmaResources.SQLVM

        if ($InTag -eq $True) {
            $ExcelSQLVM | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'ResourceGroup',
            'Name',
            'Location',
            'Zone',
            'SQL Server License Type',
            'SQL Image',
            'SQL Management',
            'SQL Image Sku',
            'Tag Name',
            'Tag Value'  | 
            Export-Excel -Path $File -WorksheetName 'SQL VMs' -AutoSize -TableName 'AzureSQLVMs' -TableStyle $tableStyle -Style $Style
        }
        else {
            $ExcelSQLVM | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'ResourceGroup',
            'Name',
            'Location',
            'Zone',
            'SQL Server License Type',
            'SQL Image',
            'SQL Management',
            'SQL Image Sku' | 
            Export-Excel -Path $File -WorksheetName 'SQL VMs' -AutoSize -TableName 'AzureSQLVMs' -TableStyle $tableStyle -Style $Style
        }
    }
    <######## Insert Column comments and documentations here following this model #########>
}