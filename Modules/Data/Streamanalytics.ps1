<######## Default Parameters. Don't modify this ########>

param($SCPath, $Sub, $Intag, $Resources, $Task , $File, $SmaResources, $TableStyle)

If ($Task -eq 'Processing') {

    <######### Insert the resource extraction here ########>

    $Streamanalytics = $Resources | Where-Object { $_.TYPE -eq 'microsoft.streamanalytics/streamingjobs' }

    $tmp = @()
    $Subs = $Sub

    foreach ($1 in $Streamanalytics) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $Creadate = (get-date $data.createdDate).ToString("yyyy-MM-dd HH:mm:ss")
        $LastOutput = (get-date $data.lastOutputEventTime).ToString("yyyy-MM-dd HH:mm:ss:ffff")
        $OutputStart = (get-date $data.outputStartTime).ToString("yyyy-MM-dd HH:mm:ss:ffff")
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
            foreach ($TagKey in $Tag.Keys) {
                $obj = @{
                    'Subscription'                      = $sub1.name;
                    'Resource Group'                    = $1.RESOURCEGROUP;
                    'Name'                              = $1.NAME;
                    'Location'                          = $1.LOCATION;
                    'SKU'                               = $data.sku.name;
                    'Compatibility Level'               = $data.compatibilityLevel;
                    'Content Storage Policy'            = $data.contentStoragePolicy;
                    'Created Date'                      = $Creadate;
                    'Data Locale'                       = $data.dataLocale;
                    'Late Arrival Max Delay in Seconds' = $data.eventsLateArrivalMaxDelayInSeconds;
                    'Out of Order Max Delay in Seconds' = $data.eventsOutOfOrderMaxDelayInSeconds;
                    'Out of Order Policy'               = $data.eventsOutOfOrderPolicy;
                    'Job State'                         = $data.jobState;
                    'Job Type'                          = $data.jobType;
                    'Last Output Event Time'            = $LastOutput;
                    'Output Start Time'                 = $OutputStart;
                    'Output Error Policy'               = $data.outputErrorPolicy;
                    'Tag Name'                          = [string]$TagKey;
                    'Tag Value'                         = [string]$Tag.$TagKey
                }
                $tmp += $obj
                if ($ResUCount -eq 1) { $ResUCount = 0 } 
            }
        }
        else {
            $obj = @{
                'Subscription'                      = $sub1.name;
                'Resource Group'                    = $1.RESOURCEGROUP;
                'Name'                              = $1.NAME;
                'Location'                          = $1.LOCATION;
                'SKU'                               = $data.sku.name;
                'Compatibility Level'               = $data.compatibilityLevel;
                'Content Storage Policy'            = $data.contentStoragePolicy;
                'Created Date'                      = $Creadate;
                'Data Locale'                       = $data.dataLocale;
                'Late Arrival Max Delay in Seconds' = $data.eventsLateArrivalMaxDelayInSeconds;
                'Out of Order Max Delay in Seconds' = $data.eventsOutOfOrderMaxDelayInSeconds;
                'Out of Order Policy'               = $data.eventsOutOfOrderPolicy;
                'Job State'                         = $data.jobState;
                'Job Type'                          = $data.jobType;
                'Last Output Event Time'            = $LastOutput;
                'Output Start Time'                 = $OutputStart;
                'Output Error Policy'               = $data.outputErrorPolicy;
                'Tag Name'                          = $null;
                'Tag Value'                         = $null
            }
            $tmp += $obj
            if ($ResUCount -eq 1) { $ResUCount = 0 } 
        }
    }
    $tmp
}
<######## Resource Excel Reporting Begins Here ########>

Else {
    <######## $SmaResources.(RESOURCE FILE NAME) ##########>

    if ($SmaResources.ExcelStreamanalytics) {
        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat 0
        $ExcelStreamanalytics = $SmaResources.ExcelStreamanalytics

        if ($InTag -eq $True) {
            $ExcelStreamanalytics | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'SKU',
            'Compatibility Level',
            'Content Storage Policy',
            'Created Date',
            'Data Locale',
            'Late Arrival Max Delay in Seconds',
            'Out of Order Max Delay in Seconds',
            'Out of Order Policy',
            'Job State',
            'Job Type',
            'Last Output Event Time',
            'Output Start Time',
            'Output Error Policy',
            'Tag Name',
            'Tag Value' | 
            Export-Excel -Path $File -WorksheetName 'Stream Analytics Jobs' -AutoSize -TableName 'AzureStreamAnalyticsJobs' -TableStyle $tableStyle -Style $Style
        }
        else {
            $ExcelStreamanalytics | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'SKU',
            'Compatibility Level',
            'Content Storage Policy',
            'Created Date',
            'Data Locale',
            'Late Arrival Max Delay in Seconds',
            'Out of Order Max Delay in Seconds',
            'Out of Order Policy',
            'Job State',
            'Job Type',
            'Last Output Event Time',
            'Output Start Time',
            'Output Error Policy' | 
            Export-Excel -Path $File -WorksheetName 'Stream Analytics Jobs' -AutoSize -TableName 'AzureStreamAnalyticsJobs' -TableStyle $tableStyle -Style $Style
        }

    }
    <######## Insert Column comments and documentations here following this model #########>
}