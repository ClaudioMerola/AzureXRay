<######## Default Parameters. Don't modify this ########>

param($SCPath, $Sub, $Intag, $Resources, $Task ,$File, $SmaResources, $TableStyle)
 
If ($Task -eq 'Processing')
{
 
    <######### Insert the resource extraction here ########>

        $vmss = $Resources | Where-Object {$_.TYPE -eq 'microsoft.compute/virtualmachinescalesets'} 

    <######### Insert the resource Process here ########>

    $tmp = @()

    $Subs = $Sub

    foreach ($1 in $vmss) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        foreach ($2 in $data.virtualMachineProfile.networkProfile.networkInterfaceConfigurations) {
            if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
                foreach ($TagKey in $Tag.Keys) {
                    $obj = @{
                        'Subscription'                  = $sub1.name;
                        'Resource Group'                = $1.RESOURCEGROUP;
                        'Name'                          = $1.NAME;
                        'Location'                      = $1.LOCATION;
                        'SKU Tier'                      = $1.sku.tier;
                        'Fault Domain'                  = $data.platformFaultDomainCount;
                        'Upgrade Policy'                = $data.upgradePolicy.mode;
                        'Capacity'                      = $1.sku.capacity;
                        'VM Size'                       = $1.sku.name;
                        'VM OS'                         = if ($null -eq $data.virtualMachineProfile.osProfile.LinuxConfiguration) { 'Windows' }else { 'Linux' };
                        'Network Interface Name'        = $2.name;
                        'Enable Accelerated Networking' = $2.properties.enableAcceleratedNetworking;
                        'Enable IP Forwarding'          = $2.properties.enableIPForwarding;
                        'Admin Username'                = $data.virtualMachineProfile.osProfile.adminUsername;
                        'VM Name Prefix'                = $data.virtualMachineProfile.osProfile.computerNamePrefix;
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
                    'Location'                      = $1.LOCATION;
                    'SKU Tier'                      = $1.sku.tier;
                    'Fault Domain'                  = $data.platformFaultDomainCount;
                    'Upgrade Policy'                = $data.upgradePolicy.mode;
                    'Capacity'                      = $1.sku.capacity;
                    'VM Size'                       = $1.sku.name;
                    'VM OS'                         = if ($null -eq $data.virtualMachineProfile.osProfile.LinuxConfiguration) { 'Windows' }else { 'Linux' };
                    'Network Interface Name'        = $2.name;
                    'Enable Accelerated Networking' = $2.properties.enableAcceleratedNetworking;
                    'Enable IP Forwarding'          = $2.properties.enableIPForwarding;
                    'Admin Username'                = $data.virtualMachineProfile.osProfile.adminUsername;
                    'VM Name Prefix'                = $data.virtualMachineProfile.osProfile.computerNamePrefix;
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

    if($SmaResources.VMSS)
    {

        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat '0'
                        
        $ExcelVMSS = $SmaResources.VMSS

        if ($InTag -eq $true) {
            $ExcelVMSS | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'SKU Tier',
            'Fault Domain',
            'Upgrade Policy',
            'Capacity',
            'VM Size',
            'VM OS',
            'Network Interface Name',
            'Enable Accelerated Networking',
            'Enable IP Forwading',
            'Admin Username',
            'VM Name Prefix',
            'Tag Name',
            'Tag Value' | 
            Export-Excel -Path $File -WorksheetName 'VMSS' -AutoSize -TableName 'AzureVMSS' -TableStyle $tableStyle -Style $Style
        }
        else {
            $ExcelVMSS | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'SKU Tier',
            'Fault Domain',
            'Upgrade Policy',
            'Capacity',
            'VM Size',
            'VM OS',
            'Network Interface Name',
            'Enable Accelerated Networking',
            'Enable IP Forwading',
            'Admin Username',
            'VM Name Prefix' | 
            Export-Excel -Path $File -WorksheetName 'VMSS' -AutoSize -TableName 'AzureVMSS' -TableStyle $tableStyle -Style $Style
        }

        <######## Insert Column comments and documentations here following this model #########>


        #$excel = Open-ExcelPackage -Path $File -KillExcel


        #Close-ExcelPackage $excel 

    }
}