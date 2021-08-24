param($SCPath, $Sub, $Intag, $Resources, $Task ,$File, $SmaResources, $TableStyle) 
If ($Task -eq 'Processing')
{

        $vm =  $Resources | Where-Object {$_.TYPE -eq 'microsoft.compute/virtualmachines'}
        $nic = $Resources | Where-Object {$_.TYPE -eq 'microsoft.network/networkinterfaces'}
        $nsg = $Resources | Where-Object {$_.TYPE -eq 'microsoft.network/networksecuritygroups'}
        $vmexp = $Resources | Where-Object {$_.TYPE -eq 'microsoft.compute/virtualmachines/extensions'}
     
    $Subs = $Sub

    $obj = ''
    $tmp = @()

    foreach ($1 in $vm) {

        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES 
        $os = if ($null -eq $data.OSProfile.LinuxConfiguration) { 'Windows' }else { 'Linux' }
        $AVSET = ''
        $dataSize = ''
        $StorAcc = ''
        $UpdateMgmt = if ($null -eq $data.osProfile.LinuxConfiguration.patchSettings.patchMode) { $data.osProfile.WindowsConfiguration.patchSettings.patchMode } else { $data.osProfile.LinuxConfiguration.patchSettings.patchMode }

        $ext = @()
        $AzDiag = ''
        $Azinsights = ''
        $ext = ($vmexp | Where-Object { ($_.id -split "/")[8] -eq $1.name }).properties.Publisher
        if ($null -ne $ext) {
        $ext = foreach ($ex in $ext) {
        if ($ex | Where-Object { $_ -eq 'Microsoft.Azure.Performance.Diagnostics' }) { $AzDiag = $true }
        if ($ex | Where-Object { $_ -eq 'Microsoft.EnterpriseCloud.Monitoring' }) { $Azinsights = $true }
        $ex + ', '
        }
        $ext = [string]$ext
        $ext = $ext.Substring(0, $ext.Length - 2)
        }
                            
        if ($null -ne $data.availabilitySet) { $AVSET = 'True' }else { $AVSET = 'False' }
        if ($data.diagnosticsProfile.bootDiagnostics.enabled -eq $true) { $bootdg = $true }else { $bootdg = $false }
        if ($null -ne $data.storageProfile.dataDisks.managedDisk.storageAccountType) {
        $StorAcc = if ($data.storageProfile.dataDisks.managedDisk.storageAccountType.count -ge 2) 
        { ($data.storageProfile.dataDisks.managedDisk.storageAccountType.count.ToString() + ' Disks found.') }
        else 
        { $data.storageProfile.dataDisks.managedDisk.storageAccountType }
        $dataSize = if ($data.storageProfile.dataDisks.managedDisk.storageAccountType.count -ge 2) 
        { ($data.storageProfile.dataDisks.diskSizeGB | Measure-Object -Sum).Sum }
        else 
        { $data.storageProfile.dataDisks.diskSizeGB }
        }

        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }

        if ($null -ne $data.networkProfile.networkInterfaces.id) {

            foreach ($2 in $data.networkProfile.networkInterfaces.id) {

                $vmnic = $nic | Where-Object { $_.ID -eq $2 }
                $vmnsg = $nsg | Where-Object { $_.properties.networkInterfaces.id -eq $2 }

                foreach ($3 in $vmnic.properties.ipConfigurations.properties) {

                    if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $True) {
                    foreach ($TagKey in $Tag.Keys) {
                        $obj = @{
                        'Subscription'                  = $sub1.name;
                        'Resource Group'                = $1.RESOURCEGROUP;
                        'Computer Name'                 = $1.NAME;
                        'Location'                      = $1.LOCATION;
                        'Zone'                          = [string]$1.ZONES;
                        'Availability Set'              = $AVSET;
                        'VM Size'                       = $data.hardwareProfile.vmSize;
                        'Image Reference'               = $data.storageProfile.imageReference.publisher;
                        'Image Version'                 = $data.storageProfile.imageReference.exactVersion;
                        'SKU'                           = $data.storageProfile.imageReference.sku;
                        'Admin Username'                = $data.osProfile.adminUsername;
                        'OS Type'                       = $os;
                        'Update Management'             = $UpdateMgmt;
                        'Boot Diagnostics'              = $bootdg;
                        'Performance Diagnostic Agent'  = if ($azDiag -ne '') { $true }else { $false };
                        'Azure Monitor'                 = if ($Azinsights -ne '') { $true }else { $false };
                        'OS Disk Storage Type'          = $data.storageProfile.osDisk.managedDisk.storageAccountType;
                        'OS Disk Size (GB)'             = $data.storageProfile.osDisk.diskSizeGB;
                        'Data Disk Storage Type'        = $StorAcc;
                        'Data Disk Size (GB)'           = $dataSize;
                        'Power State'                   = $data.extended.instanceView.powerState.displayStatus;
                        'NIC Name'                      = [string]$vmnic[0].name;
                        'NIC Type'                      = [string]$vmnic[0].properties.nicType;
                        'NSG'                           = if ($null -eq $vmnsg.NAME) { 'None' }else { $vmnsg.NAME };
                        'Enable Accelerated Networking' = [string]$vmnic[0].properties.enableAcceleratedNetworking;
                        'Enable IP Forwarding'          = [string]$vmnic[0].properties.enableIPForwarding;
                        'Primary IP'                    = $3.primary;
                        'Private IP Version'            = $3.privateIPAddressVersion;
                        'Private IP Address'            = $3.privateIPAddress;
                        'Private IP Allocation Method'  = $3.privateIPAllocationMethod;
                        'VM Extensions'                 = $ext;
                        'Resource U'                    = $ResUCount;
                        'Tag Name'                      = [string]$TagKey;
                        'Tag Value'                     = [string]$Tag.$TagKey
                        }
                        $tmp += $obj
                        if ($ResUCount -eq 1) { $ResUCount = 0 } 
                        } 
                    }
                    elseif ([string]::IsNullOrEmpty($Tag.Keys) -or $InTag -ne $True) {
                    $obj = @{
                    'Subscription'                  = $sub1.name;
                    'Resource Group'                = $1.RESOURCEGROUP;
                    'Computer Name'                 = $1.NAME;
                    'Location'                      = $1.LOCATION;
                    'Zone'                          = [string]$1.ZONES;
                    'Availability Set'              = $AVSET;
                    'VM Size'                       = $data.hardwareProfile.vmSize;
                    'Image Reference'               = $data.storageProfile.imageReference.publisher;
                    'Image Version'                 = $data.storageProfile.imageReference.exactVersion;
                    'SKU'                           = $data.storageProfile.imageReference.sku;
                    'Admin Username'                = $data.osProfile.adminUsername;
                    'OS Type'                       = $os;
                    'Update Management'             = $UpdateMgmt;
                    'Boot Diagnostics'              = $bootdg;
                    'Performance Diagnostic Agent'  = if ($azDiag -ne '') { $true }else { $false };
                    'Azure Monitor'                 = if ($Azinsights -ne '') { $true }else { $false };
                    'OS Disk Storage Type'          = $data.storageProfile.osDisk.managedDisk.storageAccountType;
                    'OS Disk Size (GB)'             = $data.storageProfile.osDisk.diskSizeGB;
                    'Data Disk Storage Type'        = $StorAcc;
                    'Data Disk Size (GB)'           = $dataSize;
                    'Power State'                   = $data.extended.instanceView.powerState.displayStatus;
                    'NIC Name'                      = [string]$vmnic[0].name;
                    'NIC Type'                      = [string]$vmnic[0].properties.nicType;
                    'NSG'                           = if ($null -eq $vmnsg.NAME) { 'None' }else { $vmnsg.NAME };
                    'Enable Accelerated Networking' = [string]$vmnic[0].properties.enableAcceleratedNetworking;
                    'Enable IP Forwarding'          = [string]$vmnic[0].properties.enableIPForwarding;
                    'Primary IP'                    = $3.primary;
                    'Private IP Version'            = $3.privateIPAddressVersion;
                    'Private IP Address'            = $3.privateIPAddress;
                    'Private IP Allocation Method'  = $3.privateIPAllocationMethod;
                    'VM Extensions'                 = $ext;
                    'Resource U'                    = $ResUCount;
                    'Tag Name'                      = $null;
                    'Tag Value'                     = $null
                    }
                    $tmp += $obj  
                    if ($ResUCount -eq 1) { $ResUCount = 0 } 
                    }   
                }
            }
        }
        elseif ($null -eq $data.networkProfile.networkInterfaces.id) {
            if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $True) {
            foreach ($TagKey in $Tag.Keys) {
                $obj = @{
                'Subscription'                  = $sub1.name;
                'Resource Group'                = $1.RESOURCEGROUP;
                'Computer Name'                 = $1.NAME;
                'Location'                      = $1.LOCATION;
                'Zone'                          = [string]$1.ZONES;
                'Availability Set'              = $AVSET;
                'VM Size'                       = $data.hardwareProfile.vmSize;
                'Image Reference'               = $data.storageProfile.imageReference.publisher;
                'Image Version'                 = $data.storageProfile.imageReference.exactVersion;
                'SKU'                           = $data.storageProfile.imageReference.sku;
                'Admin Username'                = $data.osProfile.adminUsername;
                'OS Type'                       = $os;
                'Update Management'             = $UpdateMgmt;
                'Boot Diagnostics'              = $bootdg;
                'Performance Diagnostic Agent'  = if ($azDiag -ne '') { $true }else { $false };
                'Azure Monitor'                 = if ($Azinsights -ne '') { $true }else { $false };
                'OS Disk Storage Type'          = $data.storageProfile.osDisk.managedDisk.storageAccountType;
                'OS Disk Size (GB)'             = $data.storageProfile.osDisk.diskSizeGB;
                'Data Disk Storage Type'        = $StorAcc;
                'Data Disk Size (GB)'           = $dataSize;
                'Power State'                   = $data.extended.instanceView.powerState.displayStatus;
                'NIC Name'                      = $null;
                'NIC Type'                      = $null;
                'NSG'                           = 'None';
                'Enable Accelerated Networking' = $null;
                'Enable IP Forwarding'          = $null;
                'Primary IP'                    = $null;
                'Private IP Version'            = $null;
                'Private IP Address'            = $null;
                'Private IP Allocation Method'  = $null;
                'VM Extensions'                 = $ext;
                'Resource U'                    = $ResUCount;
                'Tag Name'                      = [string]$TagKey;
                'Tag Value'                     = [string]$Tag.$TagKey
                }
                $tmp += $obj
                if ($ResUCount -eq 1) { $ResUCount = 0 } 
                }
            }
            elseif ([string]::IsNullOrEmpty($Tag.Keys) -or $InTag -ne $True) {
            $obj = @{
            'Subscription'                  = $sub1.name;
            'Resource Group'                = $1.RESOURCEGROUP;
            'Computer Name'                 = $1.NAME;
            'Location'                      = $1.LOCATION;
            'Zone'                          = [string]$1.ZONES;
            'Availability Set'              = $AVSET;
            'VM Size'                       = $data.hardwareProfile.vmSize;
            'Image Reference'               = $data.storageProfile.imageReference.publisher;
            'Image Version'                 = $data.storageProfile.imageReference.exactVersion;
            'SKU'                           = $data.storageProfile.imageReference.sku;
            'Admin Username'                = $data.osProfile.adminUsername;
            'OS Type'                       = $os;
            'Update Management'             = $UpdateMgmt;
            'Boot Diagnostics'              = $bootdg;
            'Performance Diagnostic Agent'  = if ($azDiag -ne '') { $true }else { $false };
            'Azure Monitor'                 = if ($Azinsights -ne '') { $true }else { $false };
            'OS Disk Storage Type'          = $data.storageProfile.osDisk.managedDisk.storageAccountType;
            'OS Disk Size (GB)'             = $data.storageProfile.osDisk.diskSizeGB;
            'Data Disk Storage Type'        = $StorAcc;
            'Data Disk Size (GB)'           = $dataSize;
            'Power State'                   = $data.extended.instanceView.powerState.displayStatus;
            'NIC Name'                      = $null;
            'NIC Type'                      = $null;
            'NSG'                           = 'None';
            'Enable Accelerated Networking' = $null;
            'Enable IP Forwarding'          = $null;
            'Primary IP'                    = $null;
            'Private IP Version'            = $null;
            'Private IP Address'            = $null;
            'Private IP Allocation Method'  = $null;
            'VM Extensions'                 = $ext;
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
else
{
    If($SmaResources.VM)
        {
            $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat '0' -VerticalAlignment Center
            $StyleExt = New-ExcelStyle -HorizontalAlignment Left -Range AE:AE -Width 60 -WrapText 
            $condtxtvm = $(New-ConditionalText None -Range X:X
                New-ConditionalText false -Range L:L
                New-ConditionalText falso -Range L:L
                New-ConditionalText false -Range M:M
                New-ConditionalText falso -Range M:M
                New-ConditionalText false -Range N:N
                New-ConditionalText falso -Range N:N
                New-ConditionalText false -Range Y:Y
                New-ConditionalText falso -Range Y:Y)


            $ExcelVMs = $SmaResources.VM

            if ($InTag -eq $True) {
                $ExcelVMs | 
                ForEach-Object { [PSCustomObject]$_ } | 
                Select-Object -Unique 'Subscription',
                'Resource Group',
                'Computer Name',
                'VM Size',
                'OS Type',
                'Location',
                'Image Reference',
                'Image Version',
                'SKU',
                'Admin Username',
                'Update Management',
                'Boot Diagnostics',
                'Performance Diagnostic Agent',
                'Azure Monitor',
                'OS Disk Storage Type',
                'OS Disk Size (GB)',
                'Data Disk Storage Type',
                'Data Disk Size (GB)',
                'Power State',
                'Availability Set',
                'Zone',
                'NIC Name',
                'NIC Type',
                'NSG',
                'Enable Accelerated Networking',
                'Enable IP Forwarding',
                'Primary IP',
                'Private IP Version',
                'Private IP Address',
                'Private IP Allocation Method',
                'VM Extensions',
                'Resource U',
                'Tag Name',
                'Tag Value' | 
                Export-Excel -Path $File -WorksheetName 'VMs' -TableName 'AzureVMs' -TableStyle $tableStyle -ConditionalText $condtxtvm -Style $Style, $StyleExt
            }
            else {
                $ExcelVMs | 
                ForEach-Object { [PSCustomObject]$_ } | 
                Select-Object -Unique 'Subscription',
                'Resource Group',
                'Computer Name',
                'VM Size',
                'OS Type',
                'Location',
                'Image Reference',
                'Image Version',
                'SKU',
                'Admin Username',
                'Update Management',
                'Boot Diagnostics',
                'Performance Diagnostic Agent',
                'Azure Monitor',
                'OS Disk Storage Type',
                'OS Disk Size (GB)',
                'Data Disk Storage Type',
                'Data Disk Size (GB)',
                'Power State',
                'Availability Set',
                'Zone',
                'NIC Name',
                'NIC Type',
                'NSG',
                'Enable Accelerated Networking',
                'Enable IP Forwarding',
                'Primary IP',
                'Private IP Version',
                'Private IP Address',
                'Private IP Allocation Method',
                'VM Extensions',
                'Resource U' | 
                Export-Excel -Path $File -WorksheetName 'VMs' -TableName 'AzureVMs' -TableStyle $TableStyle -ConditionalText $condtxtvm -Style $Style, $StyleExt
            }
    
            $excel = Open-ExcelPackage -Path $File -KillExcel

            $null = $excel.VMs.Cells["L1"].AddComment("Boot diagnostics is a debugging feature for Azure virtual machines (VM) that allows diagnosis of VM boot failures.", "Azure Resource Inventory")
            $excel.VMs.Cells["L1"].Hyperlink = 'https://docs.microsoft.com/en-us/azure/virtual-machines/boot-diagnostics'
            $null = $excel.VMs.Cells["M1"].AddComment("Is recommended to install Performance Diagnostics Agent in every Azure Virtual Machine upfront. The agent is only used when triggered by the console and may save time in an event of performance struggling.", "Azure Resource Inventory")
            $excel.VMs.Cells["M1"].Hyperlink = 'https://docs.microsoft.com/en-us/azure/virtual-machines/troubleshooting/performance-diagnostics'
            $null = $excel.VMs.Cells["N1"].AddComment("We recommend that you use Azure Monitor to gain visibility into your resource’s health.", "Azure Resource Inventory")
            $excel.VMs.Cells["N1"].Hyperlink = 'https://docs.microsoft.com/en-us/azure/security/fundamentals/iaas#monitor-vm-performance'
            $null = $excel.VMs.Cells["X1"].AddComment("Use a network security group to protect against unsolicited traffic into Azure subnets. Network security groups are simple, stateful packet inspection devices that use the 5-tuple approach (source IP, source port, destination IP, destination port, and layer 4 protocol) to create allow/deny rules for network traffic.", "Azure Resource Inventory")
            $excel.VMs.Cells["X1"].Hyperlink = 'https://docs.microsoft.com/en-us/azure/security/fundamentals/network-best-practices#logically-segment-subnets'
            $null = $excel.VMs.Cells["Y1"].AddComment("Accelerated networking enables single root I/O virtualization (SR-IOV) to a VM, greatly improving its networking performance. This high-performance path bypasses the host from the datapath, reducing latency, jitter, and CPU utilization.", "Azure Resource Inventory")
            $excel.VMs.Cells["Y1"].Hyperlink = 'https://docs.microsoft.com/en-us/azure/virtual-network/create-vm-accelerated-networking-cli'

            Close-ExcelPackage $excel
        }             

}