param($SCPath, $Sub, $Intag, $Resources, $Task ,$File,$SmaResources,$TableStyle) 

if ($Task -eq 'Processing')
{

$DB = @()

ForEach ($Resource in $Resources) 
    {
        if ($Resource.TYPE -eq 'microsoft.sql/servers/databases' -and $_.name -ne 'master' ) { $DB += $Resource }
    }

    $tmp = @()
    $db = $DB
    $Subs = $Sub

    foreach ($1 in $db) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $DBServer = [string]$1.id.split("/")[8]
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
            foreach ($TagKey in $Tag.Keys) {
                $obj = @{
                    'Subscription'               = $sub1.name;
                    'Resource Group'             = $1.RESOURCEGROUP;
                    'Name'                       = $1.NAME;
                    'Location'                   = $1.LOCATION;
                    'Storage Account Type'       = $data.storageAccountType;
                    'Database Server'            = $DBServer;
                    'Default Secondary Location' = $data.defaultSecondaryLocation;
                    'Status'                     = $data.status;
                    'DTU Capacity'               = $data.currentSku.capacity;
                    'DTU Tier'                   = $data.requestedServiceObjectiveName;
                    'Zone Redundant'             = $data.zoneRedundant;
                    'Catalog Collation'          = $data.catalogCollation;
                    'Read Replica Count'         = $data.readReplicaCount;
                    'Data Max Size (GB)'         = (($data.maxSizeBytes / 1024) / 1024) / 1024;
                    'Resource U'                 = $ResUCount;
                    'Tag Name'                   = [string]$TagKey;
                    'Tag Value'                  = [string]$Tag.$TagKey
                }
                $tmp += $obj
                if ($ResUCount -eq 1) { $ResUCount = 0 } 
            }
        }
        else {
            $obj = @{
                'Subscription'               = $sub1.name;
                'Resource Group'             = $1.RESOURCEGROUP;
                'Name'                       = $1.NAME;
                'Location'                   = $1.LOCATION;
                'Storage Account Type'       = $data.storageAccountType;
                'Database Server'            = $DBServer;
                'Default Secondary Location' = $data.defaultSecondaryLocation;
                'Status'                     = $data.status;
                'DTU Capacity'               = $data.currentSku.capacity;
                'DTU Tier'                   = $data.requestedServiceObjectiveName;
                'Zone Redundant'             = $data.zoneRedundant;
                'Catalog Collation'          = $data.catalogCollation;
                'Read Replica Count'         = $data.readReplicaCount;
                'Data Max Size (GB)'         = (($data.maxSizeBytes / 1024) / 1024) / 1024;
                'Resource U'                 = $ResUCount;
                'Tag Name'                   = $null;
                'Tag Value'                  = $null
            }
            $tmp += $obj
            if ($ResUCount -eq 1) { $ResUCount = 0 } 
        }
    }
    $tmp

}
else
{
    if($SmaResources.DB)
    {

        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat '0'

        $ExcelDB = $SmaResources.DB

        if ($InTag -eq $True) {
            $ExcelDB | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'Storage Account Type',
            'Database Server',
            'Default Secondary Location',
            'Status',
            'DTU Capacity',
            'DTU Tier',
            'Data Max Size (GB)',
            'Zone Redundant',
            'Catalog Collation',
            'Read Replica Count',
            'Tag Name',
            'Tag Value'  | 
            Export-Excel -Path $File -WorksheetName 'SQL DBs' -AutoSize -TableName 'AzureSQLDBs' -TableStyle $tableStyle -Style $Style
        }
        else {
            $ExcelDB | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'Storage Account Type',
            'Database Server',
            'Default Secondary Location',
            'Status',
            'DTU Capacity',
            'DTU Tier',
            'Data Max Size (GB)',
            'Zone Redundant',
            'Catalog Collation',
            'Read Replica Count' | 
            Export-Excel -Path $File -WorksheetName 'SQL DBs' -AutoSize -TableName 'AzureSQLDBs' -TableStyle $tableStyle -Style $Style
        }
    }
}