<######## Default Parameters. Don't modify this ########>

param($SCPath, $Sub, $Intag, $Resources, $Task ,$File, $SmaResources, $TableStyle)
 
If ($Task -eq 'Processing')
{
 
    <######### Insert the resource extraction here ########>

    $CONTAINER = @()

    ForEach ($Resource in $Resources) {
            if ($Resource.TYPE -eq 'microsoft.containerinstance/containergroups') { $CONTAINER += $Resource }
    }

    <######### Insert the resource Process here ########>

    $tmp = @()

    $con = $CONTAINER
    $Subs = $Sub

    foreach ($1 in $con) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        foreach ($2 in $data.containers) {
            if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
                foreach ($TagKey in $Tag.Keys) {
                    $obj = @{
                        'Subscription'        = $sub1.name;
                        'Resource Group'      = $1.RESOURCEGROUP;
                        'Instance Name'       = $1.NAME;
                        'Location'            = $1.LOCATION;
                        'Instance OS Type'    = $data.osType;
                        'Container Name'      = $2.name;
                        'Container State'     = $2.properties.instanceView.currentState.state;
                        'Container Image'     = [string]$2.properties.image;
                        'Restart Count'       = $2.properties.instanceView.restartCount;
                        'Start Time'          = $2.properties.instanceView.currentState.startTime;
                        'Command'             = [string]$2.properties.command;
                        'Request CPU'         = $2.properties.resources.requests.cpu;
                        'Request Memory (GB)' = $2.properties.resources.requests.memoryInGB;
                        'IP'                  = $data.ipAddress.ip;
                        'Protocol'            = [string]$2.properties.ports.protocol;
                        'Port'                = [string]$2.properties.ports.port;
                        'Resource U'          = $ResUCount;
                        'Tag Name'            = [string]$TagKey;
                        'Tag Value'           = [string]$Tag.$TagKey
                    }
                    $tmp += $obj
                    if ($ResUCount -eq 1) { $ResUCount = 0 } 
                }
            }
            else {
                $obj = @{
                    'Subscription'        = $sub1.name;
                    'Resource Group'      = $1.RESOURCEGROUP;
                    'Instance Name'       = $1.NAME;
                    'Location'            = $1.LOCATION;
                    'Instance OS Type'    = $data.osType;
                    'Container Name'      = $2.name;
                    'Container State'     = $2.properties.instanceView.currentState.state;
                    'Container Image'     = [string]$2.properties.image;
                    'Restart Count'       = $2.properties.instanceView.restartCount;
                    'Start Time'          = $2.properties.instanceView.currentState.startTime;
                    'Command'             = [string]$2.properties.command;
                    'Request CPU'         = $2.properties.resources.requests.cpu;
                    'Request Memory (GB)' = $2.properties.resources.requests.memoryInGB;
                    'IP'                  = $data.ipAddress.ip;
                    'Protocol'            = [string]$2.properties.ports.protocol;
                    'Port'                = [string]$2.properties.ports.port;
                    'Resource U'          = $ResUCount;
                    'Tag Name'            = $null;
                    'Tag Value'           = $null
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

    if($SmaResources.CONTAINER)
    {

        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat '0'

        $ExcelContainer = $SmaResources.CONTAINER
            
        if ($InTag -eq $true) {
            $ExcelContainer | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Instance Name',
            'Location',
            'Instance OS Type',
            'Container Name',
            'Container State',
            'Container Image',
            'Restart Count',
            'Start Time',
            'Command',
            'Request CPU',
            'Request Memory (GB)',
            'IP',
            'Protocol',
            'Port',
            'Tag Name',
            'Tag Value' | 
            Export-Excel -Path $File -WorksheetName 'Containers' -AutoSize -TableName 'AzureContainers' -TableStyle $tableStyle -Style $Style
        }
        else {
            $ExcelContainer | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Instance Name',
            'Location',
            'Instance OS Type',
            'Container Name',
            'Container State',
            'Container Image',
            'Restart Count',
            'Start Time',
            'Command',
            'Request CPU',
            'Request Memory (GB)',
            'IP',
            'Protocol',
            'Port' | 
            Export-Excel -Path $File -WorksheetName 'Containers' -AutoSize -TableName 'AzureContainers' -TableStyle $tableStyle -Style $Style
        }

        <######## Insert Column comments and documentations here following this model #########>


        #$excel = Open-ExcelPackage -Path $File -KillExcel


        #Close-ExcelPackage $excel 

    }
}