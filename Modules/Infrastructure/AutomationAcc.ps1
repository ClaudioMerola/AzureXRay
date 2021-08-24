<######## Default Parameters. Don't modify this ########>

param($SCPath, $Sub, $Intag, $Resources, $Task ,$File, $SmaResources, $TableStyle)
 
If ($Task -eq 'Processing')
{
 
    <######### Insert the resource extraction here ########>

        $runbook = $Resources | Where-Object {$_.TYPE -eq 'microsoft.automation/automationaccounts/runbooks'}
        $autacc = $Resources | Where-Object {$_.TYPE -eq 'microsoft.automation/automationaccounts'}

    <######### Insert the resource Process here ########>

    $tmp = @()
    $Subs = $Sub

    foreach ($0 in $autacc) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $0.subscriptionId }
                            
        $rbs = $runbook | Where-Object { $_.id.split('/')[8] -eq $0.name }
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        if ($null -ne $rbs) {
            foreach ($1 in $rbs) {
                if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
                    foreach ($TagKey in $Tag.Keys) {    
                        $data = $1.PROPERTIES
                        $obj = @{
                            'Subscription'             = $sub1.name;
                            'Resource Group'           = $0.RESOURCEGROUP;
                            'Automation Account Name'  = $0.NAME;
                            'Automation Account State' = $0.properties.State;
                            'Automation Account SKU'   = $0.properties.sku.name;
                            'Location'                 = $0.LOCATION;
                            'Runbook Name'             = $1.Name;
                            'Last Modified Time'       = ([datetime]$data.lastModifiedTime).tostring('MM/dd/yyyy hh:mm') ;
                            'Runbook State'            = $data.state;
                            'Runbook Type'             = $data.runbookType;
                            'Runbook Description'      = $data.description;
                            'Job Count'                = $data.jobCount;
                            'Resource U'               = $ResUCount;
                            'Tag Name'                 = [string]$TagKey;
                            'Tag Value'                = [string]$Tag.$TagKey
                        }
                        $tmp += $obj
                        if ($ResUCount -eq 1) { $ResUCount = 0 } 
                    }
                }
                else {   
                    $data = $1.PROPERTIES
                    $obj = @{
                        'Subscription'             = $sub1.name;
                        'Resource Group'           = $0.RESOURCEGROUP;
                        'Automation Account Name'  = $0.NAME;
                        'Automation Account State' = $0.properties.State;
                        'Automation Account SKU'   = $0.properties.sku.name;
                        'Location'                 = $0.LOCATION;
                        'Runbook Name'             = $1.Name;
                        'Last Modified Time'       = ([datetime]$data.lastModifiedTime).tostring('MM/dd/yyyy hh:mm') ;
                        'Runbook State'            = $data.state;
                        'Runbook Type'             = $data.runbookType;
                        'Runbook Description'      = $data.description;
                        'Job Count'                = $data.jobCount;
                        'Resource U'               = $ResUCount;
                        'Tag Name'                 = $null;
                        'Tag Value'                = $null
                    }
                    $tmp += $obj
                    if ($ResUCount -eq 1) { $ResUCount = 0 } 
                }
            }
        }
        else {
            if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
                foreach ($TagKey in $Tag.Keys) {  
                    $obj = @{
                        'Subscription'             = $sub1.name;
                        'Resource Group'           = $0.RESOURCEGROUP;
                        'Automation Account Name'  = $0.NAME;
                        'Automation Account State' = $0.properties.State;
                        'Automation Account SKU'   = $0.properties.sku.name;
                        'Location'                 = $0.LOCATION;
                        'Runbook Name'             = $null;
                        'Last Modified Time'       = $null;
                        'Runbook State'            = $null;
                        'Runbook Type'             = $null;
                        'Runbook Description'      = $null;
                        'Job Count'                = $null;
                        'Resource U'               = $ResUCount;
                        'Tag Name'                 = [string]$TagKey;
                        'Tag Value'                = [string]$Tag.$TagKey
                    }
                    $tmp += $obj
                    if ($ResUCount -eq 1) { $ResUCount = 0 } 
                }
            }
            else {   
                $obj = @{
                    'Subscription'             = $sub1.name;
                    'Resource Group'           = $0.RESOURCEGROUP;
                    'Automation Account Name'  = $0.NAME;
                    'Automation Account State' = $0.properties.State;
                    'Automation Account SKU'   = $0.properties.sku.name;
                    'Location'                 = $0.LOCATION;
                    'Runbook Name'             = $null;
                    'Last Modified Time'       = $null;
                    'Runbook State'            = $null;
                    'Runbook Type'             = $null;
                    'Runbook Description'      = $null;
                    'Job Count'                = $null;
                    'Resource U'               = $ResUCount;
                    'Tag Name'                 = $null;
                    'Tag Value'                = $null
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

    if($SmaResources.AutomationAcc)
    {

        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat '0'
        $StyleExt = New-ExcelStyle -HorizontalAlignment Left -Range K:K -Width 80 -WrapText 

        $ExcelAutAcc = $SmaResources.AutomationAcc
            
        if ($InTag -eq $true) {
            $ExcelAutAcc | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Automation Account Name',
            'Automation Account State',
            'Automation Account SKU',
            'Location',
            'Runbook Name',
            'Last Modified Time',
            'Runbook State',
            'Runbook Type',
            'Runbook Description',
            'Job Count',
            'Tag Name',
            'Tag Value' |
            Export-Excel -Path $File -WorksheetName 'Runbooks' -AutoSize -TableName 'AzureRunbooks' -TableStyle $tableStyle -Style $Style, $StyleExt
        }
        else {
            $ExcelAutAcc | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Automation Account Name',
            'Automation Account State',
            'Automation Account SKU',
            'Location',
            'Runbook Name',
            'Last Modified Time',
            'Runbook State',
            'Runbook Type',
            'Runbook Description',
            'Job Count' |
            Export-Excel -Path $File -WorksheetName 'Runbooks' -AutoSize -TableName 'AzureRunbooks' -TableStyle $tableStyle -Style $Style, $StyleExt
        }

        <######## Insert Column comments and documentations here following this model #########>


        #$excel = Open-ExcelPackage -Path $File -KillExcel


        #Close-ExcelPackage $excel 

    }
}