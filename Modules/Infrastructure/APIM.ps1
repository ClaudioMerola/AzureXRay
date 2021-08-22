<######## Default Parameters. Don't modify this ########>

param($SCPath, $Sub, $Intag, $Resources, $Task ,$File, $SmaResources, $TableStyle)
 
If ($Task -eq 'Processing')
{
 
    <######### Insert the resource extraction here ########>

    $APIM = @()

    ForEach ($Resource in $Resources) {
            if ($Resource.TYPE -eq 'microsoft.apimanagement/service' ) { $APIM += $Resource }
    }

    <######### Insert the resource Process here ########>

    $tmp = @()
    $Subs = $Sub

    foreach ($1 in $APIM) {
        $ResUCount = 1
        $sub1 = $SUBs | Where-Object { $_.id -eq $1.subscriptionId }
        $data = $1.PROPERTIES
        $Tag = @{}
        if ($data.virtualNetworkType -eq 'None') { $NetType = '' } else { $NetType = [string]$data.virtualNetworkConfiguration.subnetResourceId.split("/")[8] }
        $1.tags.psobject.properties | ForEach-Object { $Tag[$_.Name] = $_.Value }
        if (![string]::IsNullOrEmpty($Tag.Keys) -and $InTag -eq $true) {
            foreach ($TagKey in $Tag.Keys) {
                $obj = @{
                    'Subscription'         = $sub1.name;
                    'Resource Group'       = $1.RESOURCEGROUP;
                    'Name'                 = $1.NAME;
                    'Location'             = $1.LOCATION;
                    'SKU'                  = $1.sku.name;
                    'Gateway URL'          = $data.gatewayUrl;
                    'Virtual Network Type' = $data.virtualNetworkType;
                    'Virtual Network'      = $NetType;
                    'Http2'                = $data.customProperties."Microsoft.WindowsAzure.ApiManagement.Gateway.Protocols.Server.Http2";
                    'Backend SSL 3.0'      = $data.customProperties."Microsoft.WindowsAzure.ApiManagement.Gateway.Security.Backend.Protocols.Ssl30";
                    'Backend TLS 1.0'      = $data.customProperties."Microsoft.WindowsAzure.ApiManagement.Gateway.Security.Backend.Protocols.Tls10";
                    'Backend TLS 1.1'      = $data.customProperties."Microsoft.WindowsAzure.ApiManagement.Gateway.Security.Backend.Protocols.Tls11";
                    'Triple DES'           = $data.customProperties."Microsoft.WindowsAzure.ApiManagement.Gateway.Security.Ciphers.TripleDes168";
                    'Client SSL 3.0'       = $data.customProperties."Microsoft.WindowsAzure.ApiManagement.Gateway.Security.Protocols.Ssl30";
                    'Client TLS 1.0'       = $data.customProperties."Microsoft.WindowsAzure.ApiManagement.Gateway.Security.Protocols.Tls10";
                    'Client TLS 1.1'       = $data.customProperties."Microsoft.WindowsAzure.ApiManagement.Gateway.Security.Protocols.Tls11";
                    'Public IP'            = [string]$data.publicIPAddresses;
                    'Tag Name'             = [string]$TagKey;
                    'Tag Value'            = [string]$Tag.$TagKey
                }
                $tmp += $obj
                if ($ResUCount -eq 1) { $ResUCount = 0 } 
            }
        }
        else {
            $obj = @{
                'Subscription'         = $sub1.name;
                'Resource Group'       = $1.RESOURCEGROUP;
                'Name'                 = $1.NAME;
                'Location'             = $1.LOCATION;
                'SKU'                  = $1.sku.name;
                'Gateway URL'          = $data.gatewayUrl;
                'Virtual Network Type' = $data.virtualNetworkType;
                'Virtual Network'      = $NetType;
                'Http2'                = $data.customProperties."Microsoft.WindowsAzure.ApiManagement.Gateway.Protocols.Server.Http2";
                'Backend SSL 3.0'      = $data.customProperties."Microsoft.WindowsAzure.ApiManagement.Gateway.Security.Backend.Protocols.Ssl30";
                'Backend TLS 1.0'      = $data.customProperties."Microsoft.WindowsAzure.ApiManagement.Gateway.Security.Backend.Protocols.Tls10";
                'Backend TLS 1.1'      = $data.customProperties."Microsoft.WindowsAzure.ApiManagement.Gateway.Security.Backend.Protocols.Tls11";
                'Triple DES'           = $data.customProperties."Microsoft.WindowsAzure.ApiManagement.Gateway.Security.Ciphers.TripleDes168";
                'Client SSL 3.0'       = $data.customProperties."Microsoft.WindowsAzure.ApiManagement.Gateway.Security.Protocols.Ssl30";
                'Client TLS 1.0'       = $data.customProperties."Microsoft.WindowsAzure.ApiManagement.Gateway.Security.Protocols.Tls10";
                'Client TLS 1.1'       = $data.customProperties."Microsoft.WindowsAzure.ApiManagement.Gateway.Security.Protocols.Tls11";
                'Public IP'            = [string]$data.publicIPAddresses;
                'Tag Name'             = $null;
                'Tag Value'            = $null
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

    if($SmaResources.APIM)
    {

        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat '0'

        $ExcelAPIM = $SmaResources.APIM

        if ($InTag -eq $true) {
            $ExcelAPIM | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'SKU',
            'Gateway URL',
            'Virtual Network Type',
            'Virtual Network',
            'Http2',
            'Backend SSL 3.0',
            'Backend TLS 1.0',
            'Backend TLS 1.1',
            'Triple DES',
            'Client SSL 3.0',
            'Client TLS 1.0',
            'Client TLS 1.1',
            'Public IP',
            'Tag Name',
            'Tag Value' | 
            Export-Excel -Path $File -WorksheetName 'APIM' -AutoSize -TableName 'AzureAPIM' -TableStyle $tableStyle -Style $Style
        }
        else {
            $ExcelAPIM | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',
            'Resource Group',
            'Name',
            'Location',
            'SKU',
            'Gateway URL',
            'Virtual Network Type',
            'Virtual Network',
            'Http2',
            'Backend SSL 3.0',
            'Backend TLS 1.0',
            'Backend TLS 1.1',
            'Triple DES',
            'Client SSL 3.0',
            'Client TLS 1.0',
            'Client TLS 1.1',
            'Public IP' | 
            Export-Excel -Path $File -WorksheetName 'APIM' -AutoSize -TableName 'AzureAPIM' -TableStyle $tableStyle -Style $Style
        }

        <######## Insert Column comments and documentations here following this model #########>


        #$excel = Open-ExcelPackage -Path $File -KillExcel


        #Close-ExcelPackage $excel 

    }
}