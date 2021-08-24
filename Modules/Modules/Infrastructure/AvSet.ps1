<######## Default Parameters. Don't modify this ########>

param($SCPath, $Sub, $Intag, $Resources, $Task ,$File, $SmaResources, $TableStyle)
 
If ($Task -eq 'Processing')
{
 
    <######### Insert the resource extraction here ########>

        $AvSet = $Resources | Where-Object {$_.TYPE -eq 'microsoft.compute/availabilitysets'}

    <######### Insert the resource Process here ########>

    $tmp = @()
    $Subs = $Sub

    foreach ($1 in $AvSet) {
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        Foreach ($vmid in $data.virtualMachines.id) {
            $vmIds = $vmid.split('/')[8]
            if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
                foreach ($TagKey in $Tag.Keys) {
                    $obj = @{
                        'Subscription'     = $sub1.name;
                        'Resource Group'   = $1.RESOURCEGROUP;
                        'Name'             = $1.NAME;
                        'Location'         = $1.LOCATION;
                        'Fault Domains'    = [string]$data.platformFaultDomainCount;
                        'Update Domains'   = [string]$data.platformUpdateDomainCount;
                        'Virtual Machines' = [string]$vmIds;
                        'Tag Name'         = [string]$TagKey;
                        'Tag Value'        = [string]$Tag.$TagKey
                    }
                    $tmp += $obj
                }
            }
            else {
                $obj = @{
                    'Subscription'     = $sub1.name;
                    'Resource Group'   = $1.RESOURCEGROUP;
                    'Name'             = $1.NAME;
                    'Location'         = $1.LOCATION;
                    'Fault Domains'    = [string]$data.platformFaultDomainCount;
                    'Update Domains'   = [string]$data.platformUpdateDomainCount;
                    'Virtual Machines' = [string]$vmIds;
                    'Tag Name'         = $null;
                    'Tag Value'        = $null
                }
                $tmp += $obj
            }
        }
    }
    $tmp

}

<######## Resource Excel Reporting Begins Here ########>

Else
{
    <######## $SmaResources.(RESOURCE FILE NAME) ##########>

    if($SmaResources.AvSet)
    {

        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat '0'
            
        $ExcelAvSet = $SmaResources.AvSet

        if ($InTag -eq $true) {
            $ExcelAvSet | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'Fault Domains',
            'Update Domains',
            'Virtual Machines',
            'Tag Name',
            'Tag Value' | 
            Export-Excel -Path $File -WorksheetName 'Availability Sets' -AutoSize -TableName 'AvailabilitySets' -TableStyle $tableStyle -Style $Style
        }
        else {
            $ExcelAvSet | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'Fault Domains',
            'Update Domains',
            'Virtual Machines' | 
            Export-Excel -Path $File -WorksheetName 'Availability Sets' -AutoSize -TableName 'AvailabilitySets' -TableStyle $tableStyle -Style $Style
        }

        <######## Insert Column comments and documentations here following this model #########>


        #$excel = Open-ExcelPackage -Path $File -KillExcel


        #Close-ExcelPackage $excel 

    }
}