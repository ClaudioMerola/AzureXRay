<######## Default Parameters. Don't modify this ########>

param($SCPath, $Sub, $Intag, $Resources, $Task ,$File, $SmaResources, $TableStyle)
 
If ($Task -eq 'Processing')
{
 
    <######### Insert the resource extraction here ########>

        $VAULT = $Resources | Where-Object {$_.TYPE -eq 'microsoft.keyvault/vaults'}

    <######### Insert the resource Process here ########>

    $tmp = @()
    $Subs = $Sub

    foreach ($1 in $VAULT) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
            foreach ($TagKey in $Tag.Keys) {
                $obj = @{
                    'Subscription'               = $sub1.name;
                    'Resource Group'             = $1.RESOURCEGROUP;
                    'Name'                       = $1.NAME;
                    'Location'                   = $1.LOCATION;
                    'SKU Family'                 = $data.sku.family;
                    'SKU'                        = $data.sku.name;
                    'Vault Uri'                  = $data.vaultUri;
                    'Enable RBAC'                = $data.enableRbacAuthorization;
                    'Enable Soft Delete'         = $data.enableSoftDelete;
                    'Enable for Disk Encryption' = $data.enabledForDiskEncryption;
                    'Enable for Template Deploy' = $data.enabledForTemplateDeployment;
                    'Soft Delete Retention Days' = $data.softDeleteRetentionInDays;
                    'Certificate Permissions'    = [string]$data.accessPolicies.permissions.certificates;
                    'Key Permissions'            = [string]$data.accessPolicies.permissions.keys;
                    'Secret Permissions'         = [string]$data.accessPolicies.permissions.secrets;
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
                'SKU Family'                 = $data.sku.family;
                'SKU'                        = $data.sku.name;
                'Vault Uri'                  = $data.vaultUri;
                'Enable RBAC'                = $data.enableRbacAuthorization;
                'Enable Soft Delete'         = $data.enableSoftDelete;
                'Enable for Disk Encryption' = $data.enabledForDiskEncryption;
                'Enable for Template Deploy' = $data.enabledForTemplateDeployment;
                'Soft Delete Retention Days' = $data.softDeleteRetentionInDays;
                'Certificate Permissions'    = [string]$data.accessPolicies.permissions.certificates;
                'Key Permissions'            = [string]$data.accessPolicies.permissions.keys;
                'Secret Permissions'         = [string]$data.accessPolicies.permissions.secrets;
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

<######## Resource Excel Reporting Begins Here ########>

Else
{
    <######## $SmaResources.(RESOURCE FILE NAME) ##########>

    if($SmaResources.Vault)
    {

        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat '0'

        $ExcelVault = $SmaResources.Vault

        if ($InTag -eq $true) {
            $ExcelVault | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'SKU Family',
            'SKU',
            'Vault Uri',
            'Enable RBAC',
            'Enable Soft Delete',
            'Enable for Disk Encryption',
            'Enable for Template Deploy',
            'Soft Delete Retention Days',
            'Certificate Permissions',
            'Key Permissions',
            'Secret Permissions',
            'Tag Name',
            'Tag Value' | 
            Export-Excel -Path $File -WorksheetName 'Key Vaults' -AutoSize -TableName 'AzureKeyVault' -TableStyle $tableStyle -Style $Style
        }
        else {
            $ExcelVault | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'SKU Family',
            'SKU',
            'Vault Uri',
            'Enable RBAC',
            'Enable Soft Delete',
            'Enable for Disk Encryption',
            'Enable for Template Deploy',
            'Soft Delete Retention Days',
            'Certificate Permissions',
            'Key Permissions',
            'Secret Permissions' | 
            Export-Excel -Path $File -WorksheetName 'Key Vaults' -AutoSize -TableName 'AzureKeyVault' -TableStyle $tableStyle -Style $Style
        }

        <######## Insert Column comments and documentations here following this model #########>


        #$excel = Open-ExcelPackage -Path $File -KillExcel


        #Close-ExcelPackage $excel 

    }
}