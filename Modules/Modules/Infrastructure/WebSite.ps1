<######## Default Parameters. Don't modify this ########>

param($SCPath, $Sub, $Intag, $Resources, $Task ,$File, $SmaResources, $TableStyle)
 
If ($Task -eq 'Processing')
{
 
    <######### Insert the resource extraction here ########>

        $WebSite = $Resources | Where-Object {$_.TYPE -eq 'microsoft.web/sites'}

    <######### Insert the resource Process here ########>

    $tmp = @()
    $Subs = $Sub

    foreach ($1 in $WebSite) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        foreach ($2 in $data.hostNameSslStates) {
            if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
                foreach ($TagKey in $Tag.Keys) {
                    $obj = @{
                        'Subscription'                  = $sub1.name;
                        'Resource Group'                = $1.RESOURCEGROUP;
                        'Name'                          = $1.NAME;
                        'Kind'                          = $1.KIND;
                        'Location'                      = $1.LOCATION;
                        'Enabled'                       = $data.enabled;
                        'state'                         = $data.state;
                        'SKU'                           = $data.sku;
                        'Content Availability State'    = $data.contentAvailabilityState;
                        'Runtime Availability State'    = $data.runtimeAvailabilityState;
                        'Possible Inbound IP Addresses' = $data.possibleInboundIpAddresses;
                        'Repository Site Name'          = $data.repositorySiteName;
                        'Availability State'            = $data.availabilityState;
                        'HostNames'                     = $2.Name;
                        'HostName Type'                 = $2.hostType;
                        'ssl State'                     = $2.sslState;
                        'Default Hostname'              = $data.defaultHostName;
                        'Client Cert Mode'              = $data.clientCertMode;
                        'ContainerSize'                 = $data.containerSize;
                        'Admin Enabled'                 = $data.adminEnabled;
                        'FTPs Host Name'                = $data.ftpsHostName;
                        'HTTPS Only'                    = $data.httpsOnly;
                        'Resource U'                    = $ResUCount;
                        'Tag Name'                      = [string]$TagKey;
                        'Tag Value'                     = [string]$Tag.$TagKey
                    }
                    $tmp += $obj
                    if ($ResUCount -eq 1) { $ResUCount = 0 } 
                }
            }
            else {
                $obj = @{
                    'Subscription'                  = $sub1.name;
                    'Resource Group'                = $1.RESOURCEGROUP;
                    'Name'                          = $1.NAME;
                    'Kind'                          = $1.KIND;
                    'Location'                      = $1.LOCATION;
                    'Enabled'                       = $data.enabled;
                    'state'                         = $data.state;
                    'SKU'                           = $data.sku;
                    'Content Availability State'    = $data.contentAvailabilityState;
                    'Runtime Availability State'    = $data.runtimeAvailabilityState;
                    'Possible Inbound IP Addresses' = $data.possibleInboundIpAddresses;
                    'Repository Site Name'          = $data.repositorySiteName;
                    'Availability State'            = $data.availabilityState;
                    'HostNames'                     = $2.Name;
                    'HostName Type'                 = $2.hostType;
                    'ssl State'                     = $2.sslState;
                    'Default Hostname'              = $data.defaultHostName;
                    'Client Cert Mode'              = $data.clientCertMode;
                    'ContainerSize'                 = $data.containerSize;
                    'Admin Enabled'                 = $data.adminEnabled;
                    'FTPs Host Name'                = $data.ftpsHostName;
                    'HTTPS Only'                    = $data.httpsOnly;
                    'Resource U'                    = $ResUCount;
                    'Tag Name'                      = $null;
                    'Tag Value'                     = $null
                }
                $tmp += $obj
                if ($ResUCount -eq 1) { $ResUCount = 0 } 
            }
        }
    }
    $tmp

}

<######## Resource Excel Reporting Begins Here ########>

Else
{
    <######## $SmaResources.(RESOURCE FILE NAME) ##########>

    if($SmaResources.WebSite)
    {

        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat '0'

        $ExcelWebSite = $SmaResources.WebSite

        if ($InTag -eq $true) {
            $ExcelWebSite | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Kind',
            'Location',
            'Enabled',
            'State',
            'SKU',
            'Content Availability State',
            'Runtime Availability State',
            'Possible Inbound IP Addresses',
            'Repository Site Name',
            'AvailabilityState',
            'HostNames',
            'HostName Type',
            'sslState',
            'Default Hostname',
            'Client Cert Mode',
            'ContainerSize',
            'Admin Enabled',
            'FTPs Host Name',
            'HTTPS Only',
            'Tag Name',
            'Tag Value' | 
            Export-Excel -Path $File -WorksheetName 'Web Sites' -AutoSize -TableName 'WebSites' -TableStyle $tableStyle -Style $Style
        }
        else {
            $ExcelWebSite | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Kind',
            'Location',
            'Enabled',
            'State',
            'SKU',
            'Content Availability State',
            'Runtime Availability State',
            'Possible Inbound IP Addresses',
            'Repository Site Name',
            'AvailabilityState',
            'HostNames',
            'HostName Type',
            'sslState',
            'Default Hostname',
            'Client Cert Mode',
            'ContainerSize',
            'Admin Enabled',
            'FTPs Host Name',
            'HTTPS Only' | 
            Export-Excel -Path $File -WorksheetName 'Web Sites' -AutoSize -TableName 'WebSites' -TableStyle $tableStyle -Style $Style
        }

        <######## Insert Column comments and documentations here following this model #########>


        #$excel = Open-ExcelPackage -Path $File -KillExcel


        #Close-ExcelPackage $excel 

    }
}