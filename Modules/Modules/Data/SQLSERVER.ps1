param($SCPath, $Sub, $Intag, $Resources, $Task , $File, $SmaResources, $TableStyle) 

if ($Task -eq 'Processing') {

    $SQLSERVER = $Resources | Where-Object { $_.TYPE -eq 'microsoft.sql/servers' }

    $tmp = @()
    $Subs = $Sub

    foreach ($1 in $SQLSERVER) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
            foreach ($TagKey in $Tag.Keys) {
                $obj = @{
                    'Subscription'          = $sub1.name;
                    'Resource Group'        = $1.RESOURCEGROUP;
                    'Name'                  = $1.NAME;
                    'Location'              = $1.LOCATION;
                    'Kind'                  = $1.kind;
                    'Admin Login'           = $data.administratorLogin;
                    'FQDN'                  = $data.fullyQualifiedDomainName;
                    'Public Network Access' = $data.publicNetworkAccess;
                    'State'                 = $data.state;
                    'Version'               = $data.version;
                    'Resource U'            = $ResUCount;
                    'Zone Redundant'        = $1.zones;
                    'Tag Name'              = [string]$TagKey;
                    'Tag Value'             = [string]$Tag.$TagKey
                }
                $tmp += $obj
                if ($ResUCount -eq 1) { $ResUCount = 0 } 
            }
        }
        else {
            $obj = @{
                'Subscription'          = $sub1.name;
                'Resource Group'        = $1.RESOURCEGROUP;
                'Name'                  = $1.NAME;
                'Location'              = $1.LOCATION;
                'Kind'                  = $1.kind;
                'Admin Login'           = $data.administratorLogin;
                'FQDN'                  = $data.fullyQualifiedDomainName;
                'Public Network Access' = $data.publicNetworkAccess;
                'State'                 = $data.state;
                'Version'               = $data.version;
                'Resource U'            = $ResUCount;
                'Zone Redundant'        = $1.zones;
                'Tag Name'              = $null;
                'Tag Value'             = $null
            }
            $tmp += $obj
            if ($ResUCount -eq 1) { $ResUCount = 0 } 
        }
    }
    $tmp
}
else {
    if ($SmaResources.SQLSERVER) {
        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat 0
        $ExcelSQLServer = $SmaResources.SQLSERVER

        if ($InTag -eq $True) {
            $ExcelSQLServer | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'Kind',
            'Admin Login',
            'FQDN',
            'Public Network Access',
            'State',
            'Version',
            'Resource U',
            'Zone Redundant',
            'Tag Name',
            'Tag Value'  | 
            Export-Excel -Path $File -WorksheetName 'SQL Servers' -AutoSize -TableName 'AzureSQLServers' -TableStyle $tableStyle -Style $Style
        }
        else {
            $ExcelSQLServer | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'Kind',
            'Admin Login',
            'FQDN',
            'Public Network Access',
            'State',
            'Version',
            'Resource U',
            'Zone Redundant' | 
            Export-Excel -Path $File -WorksheetName 'SQL Servers' -AutoSize -TableName 'AzureSQLServers' -TableStyle $tableStyle -Style $Style
        }
    }
}