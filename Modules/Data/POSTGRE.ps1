<######## Default Parameters. Don't modify this ########>

param($SCPath, $Sub, $Intag, $Resources, $Task , $File, $SmaResources, $TableStyle)

If ($Task -eq 'Processing') {

    $POSTGRE = $Resources | Where-Object { $_.TYPE -eq 'microsoft.dbforpostgresql/servers' }

    $tmp = @()
    $Subs = $Sub


    foreach ($1 in $POSTGRE) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $sku = $1.SKU
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
            foreach ($TagKey in $Tag.Keys) {
                $obj = @{
                    'Subscription'              = $sub1.name;
                    'Resource Group'            = $1.RESOURCEGROUP;
                    'Name'                      = $1.NAME;
                    'Location'                  = $1.LOCATION;
                    'SKU'                       = $sku.name;
                    'SKU Family'                = $sku.family;
                    'Tier'                      = $sku.tier;
                    'Capacity'                  = $sku.capacity;
                    'Postgre Version'           = $data.version;
                    'Backup Retention Days'     = $data.storageProfile.backupRetentionDays;
                    'Geo-Redundant Backup'      = $data.storageProfile.geoRedundantBackup;
                    'Auto Grow'                 = $data.storageProfile.storageAutogrow;
                    'Storage MB'                = $data.storageProfile.storageMB;
                    'Public Network Access'     = $data.publicNetworkAccess;
                    'Admin Login'               = $data.administratorLogin;
                    'Infrastructure Encryption' = $data.InfrastructureEncryption;
                    'Minimal Tls Version'       = $data.minimalTlsVersion;
                    'State'                     = $data.userVisibleState;
                    'Replica Capacity'          = $data.replicaCapacity;
                    'Replication Role'          = $data.replicationRole;
                    'BYOK Enforcement'          = $data.byokEnforcement;
                    'ssl Enforcement'           = $data.sslEnforcement;
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
                'SKU'                       = $sku.name;
                'SKU Family'                = $sku.family;
                'Tier'                      = $sku.tier;
                'Capacity'                  = $sku.capacity;
                'Postgre Version'           = $data.version;
                'Backup Retention Days'     = $data.storageProfile.backupRetentionDays;
                'Geo-Redundant Backup'      = $data.storageProfile.geoRedundantBackup;
                'Auto Grow'                 = $data.storageProfile.storageAutogrow;
                'Storage MB'                = $data.storageProfile.storageMB;
                'Public Network Access'     = $data.publicNetworkAccess;
                'Admin Login'               = $data.administratorLogin;
                'Infrastructure Encryption' = $data.InfrastructureEncryption;
                'Minimal Tls Version'       = $data.minimalTlsVersion;
                'State'                     = $data.userVisibleState;
                'Replica Capacity'          = $data.replicaCapacity;
                'Replication Role'          = $data.replicationRole;
                'BYOK Enforcement'          = $data.byokEnforcement;
                'ssl Enforcement'           = $data.sslEnforcement;
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
<######## Resource Excel Reporting Begins Here ########>

Else {
    <######## $SmaResources.(RESOURCE FILE NAME) ##########>

    if ($SmaResources.POSTGRE) {
        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat 0
        $ExcelPOSTGRE = $SmaResources.POSTGRE

        if ($InTag -eq $True) {
            $ExcelPostGre | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'SKU',
            'SKU Family',
            'Tier',
            'Capacity',
            'Postgre Version',
            'Backup Retention Days',
            'Geo-Redundant Backup',
            'Auto Grow',
            'Storage MB',
            'Public Network Access',
            'Admin Login',
            'Infrastructure Encryption',
            'Minimal Tls Version',
            'State',
            'Replica Capacity',
            'Replication Role',
            'BYOK Enforcement',
            'ssl Enforcement',
            'Tag Name',
            'Tag Value' | 
            Export-Excel -Path $File -WorksheetName 'PostgreSQL' -AutoSize -TableName 'AzurePostgreSQL' -TableStyle $tableStyle -Style $Style
        }
        else {
            $ExcelPostGre | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'SKU',
            'SKU Family',
            'Tier',
            'Capacity',
            'Postgre Version',
            'Backup Retention Days',
            'Geo-Redundant Backup',
            'Auto Grow',
            'Storage MB',
            'Public Network Access',
            'Admin Login',
            'Infrastructure Encryption',
            'Minimal Tls Version',
            'State',
            'Replica Capacity',
            'Replication Role',
            'BYOK Enforcement',
            'ssl Enforcement' | 
            Export-Excel -Path $File -WorksheetName 'PostgreSQL' -AutoSize -TableName 'AzurePostgreSQL' -TableStyle $tableStyle -Style $Style
        }

    }
    <######## Insert Column comments and documentations here following this model #########>
}