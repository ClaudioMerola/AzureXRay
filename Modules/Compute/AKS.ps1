<######## Default Parameters. Don't modify this ########>

param($SCPath, $Sub, $Intag, $Resources, $Task ,$File, $SmaResources, $TableStyle)
 
If ($Task -eq 'Processing')
{
 
    <######### Insert the resource extraction here ########>

    $AKS = @()

    ForEach ($Resource in $Resources) {
            if ($Resource.TYPE -eq 'microsoft.containerservice/managedclusters' ) { $AKS += $Resource }
    }

    <######### Insert the resource Process here ########>

    $tmp = @()

    $AKS = $AKS
    $Subs = $Sub

    foreach ($1 in $AKS) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        if ($data.kubernetesVersion -lt 1.17) {
            $ver = 'UNSUPPORTED'
        }
        else {
            $ver = 'SUPPORTED'
        }
        $Tag = @{}
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        foreach ($2 in $data.agentPoolProfiles) {
            if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
                foreach ($TagKey in $Tag.Keys) {
                    $obj = @{
                        'Subscription'               = $sub1.name;
                        'Resource Group'             = $1.RESOURCEGROUP;
                        'Clusters'                   = $1.NAME;
                        'Location'                   = $1.LOCATION;
                        'Kubernetes Version'         = $data.kubernetesVersion;
                        'Kubernetes Version Support' = $ver;
                        'Role-Based Access Control'  = $data.enableRBAC;
                        'AAD Enabled'                = if ($data.aadProfile) { $true }else { $false };
                        'Network Type'               = $data.networkProfile.networkPlugin;
                        'Outbound Type'              = $data.networkProfile.outboundType;
                        'LoadBalancer Sku'           = $data.networkProfile.loadBalancerSku;
                        'Docker Pod Cidr'            = $data.networkProfile.podCidr;
                        'Service Cidr'               = $data.networkProfile.serviceCidr;
                        'Docker Bridge Cidr'         = $data.networkProfile.dockerBridgeCidr;                   
                        'Network DNS Service IP'     = $data.networkProfile.dnsServiceIP;
                        'FQDN'                       = $data.fqdn
                        'HTTP Application Routing'   = if ($data.addonProfiles.httpapplicationrouting.enabled) { $true }else { $false };
                        'Node Pool Name'             = $2.name;
                        'Pool Profile Type'          = $2.type;
                        'Pool OS'                    = $2.osType;
                        'Node Size'                  = $2.vmSize;
                        'OS Disk Size (GB)'          = $2.osDiskSizeGB;
                        'Nodes'                      = $2.count;
                        'Autoscale'                  = $2.enableAutoScaling;
                        'Autoscale Max'              = $2.maxCount;
                        'Autoscale Min'              = $2.minCount;
                        'Max Pods Per Node'          = $2.maxPods;
                        'Orchestrator Version'       = $2.orchestratorVersion;
                        'Enable Node Public IP'      = $2.enableNodePublicIP;
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
                    'Clusters'                   = $1.NAME;
                    'Location'                   = $1.LOCATION;
                    'Kubernetes Version'         = $data.kubernetesVersion;
                    'Kubernetes Version Support' = $ver;
                    'Role-Based Access Control'  = $data.enableRBAC;
                    'AAD Enabled'                = if ($data.aadProfile) { $true }else { $false };
                    'Network Type'               = $data.networkProfile.networkPlugin;
                    'Outbound Type'              = $data.networkProfile.outboundType;
                    'LoadBalancer Sku'           = $data.networkProfile.loadBalancerSku;
                    'Docker Pod Cidr'            = $data.networkProfile.podCidr;
                    'Service Cidr'               = $data.networkProfile.serviceCidr;
                    'Docker Bridge Cidr'         = $data.networkProfile.dockerBridgeCidr;                   
                    'Network DNS Service IP'     = $data.networkProfile.dnsServiceIP;
                    'FQDN'                       = $data.fqdn
                    'HTTP Application Routing'   = if ($data.addonProfiles.httpapplicationrouting.enabled) { $true }else { $false };
                    'Node Pool Name'             = $2.name;
                    'Pool Profile Type'          = $2.type;
                    'Pool OS'                    = $2.osType;
                    'Node Size'                  = $2.vmSize;
                    'OS Disk Size (GB)'          = $2.osDiskSizeGB;
                    'Nodes'                      = $2.count;
                    'Autoscale'                  = $2.enableAutoScaling;
                    'Autoscale Max'              = $2.maxCount;
                    'Autoscale Min'              = $2.minCount;
                    'Max Pods Per Node'          = $2.maxPods;
                    'Orchestrator Version'       = $2.orchestratorVersion;
                    'Enable Node Public IP'      = $2.enableNodePublicIP;
                    'Resource U'                 = $ResUCount;
                    'Tag Name'                   = $null;
                    'Tag Value'                  = $null
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

    if($SmaResources.AKS)
    {

        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat '0'

        $txtaksv = New-ConditionalText UNSUPPORTED -Range F:F

        $ExcelAKS = $SmaResources.AKS

        if ($InTag -eq $true) {
            $ExcelAKS | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Clusters',
            'Location',
            'Kubernetes Version',
            'Kubernetes Version Support',
            'Role-Based Access Control',
            'AAD Enabled',
            'Network Type',
            'Outbound Type',
            'LoadBalancer Sku',
            'Docker Pod Cidr',
            'Service Cidr',
            'Docker Bridge Cidr',           
            'Network DNS Service IP',
            'FQDN',
            'HTTP Application Routing',
            'Node Pool Name',
            'Pool Profile Type',
            'Pool OS',
            'Node Size',
            'OS Disk Size (GB)',
            'Nodes',
            'Autoscale',
            'Autoscale Max',
            'Autoscale Min',
            'Max Pods Per Node',
            'Orchestrator Version',
            'Enable Node Public IP',
            'Tag Name',
            'Tag Value' | 
            Export-Excel -Path $File -WorksheetName 'AKS' -AutoSize -TableName 'AzureKubernetes' -TableStyle $tableStyle -ConditionalText $txtaksv -Numberformat '0' -Style $Style
        }
        else {
            $ExcelAKS | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Clusters',
            'Location',
            'Kubernetes Version',
            'Kubernetes Version Support',
            'Role-Based Access Control',
            'AAD Enabled',
            'Network Type',
            'Outbound Type',
            'LoadBalancer Sku',
            'Docker Pod Cidr',
            'Service Cidr',
            'Docker Bridge Cidr',           
            'Network DNS Service IP',
            'FQDN',
            'HTTP Application Routing',
            'Node Pool Name',
            'Pool Profile Type',
            'Pool OS',
            'Node Size',
            'OS Disk Size (GB)',
            'Nodes',
            'Autoscale',
            'Autoscale Max',
            'Autoscale Min',
            'Max Pods Per Node',
            'Orchestrator Version',
            'Enable Node Public IP' | 
            Export-Excel -Path $File -WorksheetName 'AKS' -AutoSize -TableName 'AzureKubernetes' -TableStyle $tableStyle -ConditionalText $txtaksv -Numberformat '0' -Style $Style
        }

        <######## Insert Column comments and documentations here following this model #########>


        #$excel = Open-ExcelPackage -Path $File -KillExcel


        #Close-ExcelPackage $excel 

    }
}