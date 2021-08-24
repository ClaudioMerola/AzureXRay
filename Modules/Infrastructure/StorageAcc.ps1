<######## Default Parameters. Don't modify this ########>

param($SCPath, $Sub, $Intag, $Resources, $Task ,$File, $SmaResources, $TableStyle)
 
If ($Task -eq 'Processing')
{
 
    <######### Insert the resource extraction here ########>

        $storageacc = $Resources | Where-Object {$_.TYPE -eq 'microsoft.storage/storageaccounts'}

    <######### Insert the resource Process here ########>

    $tmp = @()
    $Subs = $Sub

    foreach ($1 in $storageacc) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $TLSv = if ($data.minimumTlsVersion -eq 'TLS1_2') { "TLS 1.2" }elseif ($data.minimumTlsVersion -eq 'TLS1_1') { "TLS 1.1" }else { "TLS 1.0" }
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
            foreach ($TagKey in $Tag.Keys) {   
                $obj = @{
                    'Subscription'                          = $sub1.name;
                    'Resource Group'                        = $1.RESOURCEGROUP;
                    'Name'                                  = $1.NAME;
                    'Location'                              = $1.LOCATION;
                    'Zone'                                  = $1.ZONES;
                    'Supports HTTPs Traffic Only'           = $data.supportsHttpsTrafficOnly;
                    'Allow Blob Public Access'              = if ($data.allowBlobPublicAccess -eq $false) { $false }else { $true };
                    'TLS Version'                           = $TLSv;
                    'Identity-based access for file shares' = if ($data.azureFilesIdentityBasedAuthentication.directoryServiceOptions -eq 'None') { $false }elseif ($null -eq $data.azureFilesIdentityBasedAuthentication.directoryServiceOptions) { $false }else { $true };
                    'Access Tier'                           = $data.accessTier;
                    'Primary Location'                      = $data.primaryLocation;
                    'Status Of Primary'                     = $data.statusOfPrimary;
                    'Secondary Location'                    = $data.secondaryLocation;
                    'Blob Address'                          = [string]$data.primaryEndpoints.blob;
                    'File Address'                          = [string]$data.primaryEndpoints.file;
                    'Table Address'                         = [string]$data.primaryEndpoints.table;
                    'Queue Address'                         = [string]$data.primaryEndpoints.queue;
                    'Network Acls'                          = $data.networkAcls.defaultAction;
                    'Resource U'                            = $ResUCount;
                    'Tag Name'                              = [string]$TagKey;
                    'Tag Value'                             = [string]$Tag.$TagKey
                }
                $tmp += $obj
                if ($ResUCount -eq 1) { $ResUCount = 0 } 
            }
        }
        else {   
            $obj = @{
                'Subscription'                          = $sub1.name;
                'Resource Group'                        = $1.RESOURCEGROUP;
                'Name'                                  = $1.NAME;
                'Location'                              = $1.LOCATION;
                'Zone'                                  = $1.ZONES;
                'Supports HTTPs Traffic Only'           = $data.supportsHttpsTrafficOnly;
                'Allow Blob Public Access'              = if ($data.allowBlobPublicAccess -eq $false) { $false }else { $true };
                'TLS Version'                           = $TLSv;
                'Identity-based access for file shares' = if ($data.azureFilesIdentityBasedAuthentication.directoryServiceOptions -eq 'None') { $false }elseif ($null -eq $data.azureFilesIdentityBasedAuthentication.directoryServiceOptions) { $false }else { $true };
                'Access Tier'                           = $data.accessTier;
                'Primary Location'                      = $data.primaryLocation;
                'Status Of Primary'                     = $data.statusOfPrimary;
                'Secondary Location'                    = $data.secondaryLocation;
                'Blob Address'                          = [string]$data.primaryEndpoints.blob;
                'File Address'                          = [string]$data.primaryEndpoints.file;
                'Table Address'                         = [string]$data.primaryEndpoints.table;
                'Queue Address'                         = [string]$data.primaryEndpoints.queue;
                'Network Acls'                          = $data.networkAcls.defaultAction;
                'Resource U'                            = $ResUCount;
                'Tag Name'                              = $null;
                'Tag Value'                             = $null
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

    if($SmaResources.StorageAcc)
    {
        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat 0

        $condtxtStorage = $(New-ConditionalText false -Range F:F
            New-ConditionalText falso -Range F:F
            New-ConditionalText true -Range G:G
            New-ConditionalText verdadeiro -Range G:G
            New-ConditionalText 1.0 -Range H:H)

        $ExcelStorageAcc = $SmaResources.StorageAcc
            
        if ($InTag -eq $True) {
            $ExcelStorageAcc | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'Zone',
            'Supports HTTPS Traffic Only',
            'Allow Blob Public Access',
            'TLS Version',
            'Identity-based access for file shares',
            'Access Tier',
            'Primary Location',
            'Status Of Primary',
            'Secondary Location',
            'Blob Address',
            'File Address',
            'Table Address',
            'Queue Address',
            'Network Acls',
            'Tag Name',
            'Tag Value'  | 
            Export-Excel -Path $File -WorksheetName 'StorageAcc' -AutoSize -TableName 'AzureStorageAccs' -TableStyle $tableStyle -ConditionalText $condtxtStorage -Style $Style
        }
        else {
            $ExcelStorageAcc | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'Zone',
            'Supports HTTPS Traffic Only',
            'Allow Blob Public Access',
            'TLS Version',
            'Identity-based access for file shares',
            'Access Tier',
            'Primary Location',
            'Status Of Primary',
            'Secondary Location',
            'Blob Address',
            'File Address',
            'Table Address',
            'Queue Address',
            'Network Acls' | 
            Export-Excel -Path $File -WorksheetName 'StorageAcc' -AutoSize -TableName 'AzureStorageAccs' -TableStyle $tableStyle -ConditionalText $condtxtStorage -Style $Style
        }


        <######## Insert Column comments and documentations here following this model #########>


        $excel = Open-ExcelPackage -Path $File -KillExcel

        $null = $excel.StorageAcc.Cells["F1"].AddComment("Is recommended that you configure your storage account to accept requests from secure connections only by setting the Secure transfer required property for the storage account.", "Azure Resource Inventory")
        $excel.StorageAcc.Cells["F1"].Hyperlink = 'https://docs.microsoft.com/en-us/azure/storage/common/storage-require-secure-transfer'
        $null = $excel.StorageAcc.Cells["G1"].AddComment("When a container is configured for public access, any client can read data in that container. Public access presents a potential security risk, so if your scenario does not require it, Microsoft recommends that you disallow it for the storage account.", "Azure Resource Inventory")
        $excel.StorageAcc.Cells["G1"].Hyperlink = 'https://docs.microsoft.com/en-us/azure/storage/blobs/anonymous-read-access-configure?tabs=portal'
        $null = $excel.StorageAcc.Cells["H1"].AddComment("By default, Azure Storage accounts permit clients to send and receive data with the oldest version of TLS, TLS 1.0, and above. To enforce stricter security measures, you can configure your storage account to require that clients send and receive data with a newer version of TLS", "Azure Resource Inventory")
        $excel.StorageAcc.Cells["H1"].Hyperlink = 'https://docs.microsoft.com/en-us/azure/storage/common/transport-layer-security-configure-minimum-version?tabs=portal'

        Close-ExcelPackage $excel 

    }
}