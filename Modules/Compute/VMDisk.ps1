<######## Default Parameters. Don't modify this ########>

param($SCPath, $Sub, $Intag, $Resources, $Task ,$File, $SmaResources, $TableStyle)
 
If ($Task -eq 'Processing')
{
 
    <######### Insert the resource extraction here ########>

        $disk = $Resources | Where-Object {$_.TYPE -eq 'microsoft.compute/disks'}

    <######### Insert the resource Process here ########>

    $tmp = @()

    $Subs = $Sub

    foreach ($1 in $disk) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $SKU = $1.SKU 
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
            foreach ($TagKey in $Tag.Keys) {
                $obj = @{
                    'Subscription'           = $sub1.name;
                    'Resource Group'         = $1.RESOURCEGROUP;
                    'Virtual Machine'        = $1.MANAGEDBY.split('/')[8];
                    'Disk Name'              = $1.NAME;
                    'Location'               = $1.LOCATION;
                    'Zone'                   = [string]$1.ZONES;
                    'SKU'                    = $SKU.Name;
                    'Disk Size'              = $data.diskSizeGB;
                    'Encryption'             = $data.encryption.type;
                    'OS Type'                = $data.osType;
                    'Disk IOPS Read / Write' = $data.diskIOPSReadWrite;
                    'Disk MBps Read / Write' = $data.diskMBpsReadWrite;
                    'Disk State'             = $data.diskState;
                    'HyperV Generation'      = $data.hyperVGeneration;
                    'Resource U'             = $ResUCount;
                    'Tag Name'               = [string]$TagKey;
                    'Tag Value'              = [string]$Tag.$TagKey
                }         
                $tmp += $obj
                if ($ResUCount -eq 1) { $ResUCount = 0 } 
            }
        }
        else {
            $obj = @{
                'Subscription'           = $sub1.name;
                'Resource Group'         = $1.RESOURCEGROUP;
                'Virtual Machine'        = $1.MANAGEDBY.split('/')[8];
                'Disk Name'              = $1.NAME;
                'Location'               = $1.LOCATION;
                'Zone'                   = [string]$1.ZONES;
                'SKU'                    = $SKU.Name;
                'Disk Size'              = $data.diskSizeGB;
                'Encryption'             = $data.encryption.type;
                'OS Type'                = $data.osType;
                'Disk IOPS Read / Write' = $data.diskIOPSReadWrite;
                'Disk MBps Read / Write' = $data.diskMBpsReadWrite;
                'Disk State'             = $data.diskState;
                'HyperV Generation'      = $data.hyperVGeneration;
                'Resource U'             = $ResUCount;
                'Tag Name'               = $null;
                'Tag Value'              = $null
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

    if($SmaResources.VMDisk)
    {

        $condtxtdsk = New-ConditionalText Unattached -Range K:K
        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat '0'
         

        $ExcelVMDisks = $SmaResources.VMDisk
                        
        if ($InTag -eq $True) {
            $ExcelVMDisks | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Virtual Machine',
            'Disk Name',
            'Zone',
            'SKU',
            'Disk Size',
            'Location',
            'Encryption',
            'OS Type',
            'Disk State',
            'Disk IOPS Read / Write',
            'Disk MBps Read / Write',
            'HyperV Generation',
            'Tag Name',
            'Tag Value' | 
            Export-Excel -Path $File -WorksheetName 'Disks' -TableName 'AzureDisks' -TableStyle $tableStyle -ConditionalText $condtxtdsk -Style $Style
        }
        else {
            $ExcelVMDisks | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Virtual Machine',
            'Disk Name',
            'Zone',
            'SKU',
            'Disk Size',
            'Location',
            'Encryption',
            'OS Type',
            'Disk State',
            'Disk IOPS Read / Write',
            'Disk MBps Read / Write',
            'HyperV Generation' | 
            Export-Excel -Path $File -WorksheetName 'Disks' -TableName 'AzureDisks' -TableStyle $tableStyle -ConditionalText $condtxtdsk -Style $Style
        }

        <######## Insert Column comments and documentations here following this model #########>

        #$excel = Open-ExcelPackage -Path $File -KillExcel


        #Close-ExcelPackage $excel 

    }
}