<######## Default Parameters. Don't modify this ########>

param($Subscriptions, $Resources, $DFile)

Function Variables0 
{

    $Global:AZVGWs = $resources | where {$_.Type -eq 'microsoft.network/virtualnetworkgateways'} | Select-Object -Property * -Unique
    $Global:AZLGWs = $resources | where {$_.Type -eq 'microsoft.network/localnetworkgateways'} | Select-Object -Property * -Unique
    $Global:AZVNETs = $resources | where {$_.Type -eq 'microsoft.network/virtualnetworks'} | Select-Object -Property * -Unique
    $Global:AZCONs = $resources | where {$_.Type -eq 'microsoft.network/connections'} | Select-Object -Property * -Unique
    $Global:AZEXPROUTEs = $resources | where {$_.Type -eq 'microsoft.network/expressroutecircuits'} | Select-Object -Property * -Unique

    $Global:Diag = Get-ChildItem -Path "C:\Program Files\" -Name "AZUREDIAGRAMS_M.VSTX"  -Recurse
    if(!$Global:Diag)
    {
        $Global:Diag = Get-ChildItem -Path "C:\Program Files (x86)\" -Name "AZUREDIAGRAMS_M.VSTX"  -Recurse
        $Global:Path = ('C:\Program Files (x86)\'+$Diag.Replace("AZUREDIAGRAMS_M.VSTX",""))
    }
    else
    {
        $Global:Path = ('C:\Program Files\'+$Diag.Replace("AZUREDIAGRAMS_M.VSTX",""))
    }

    $Global:AzNet = ($Global:Path+"AZURENETWORKING_M.VSSX")
    $Global:AzGen = ($Global:Path+"AZUREGENERAL_M.VSSX")
    $Global:AzCom = ($Global:Path+"AZURECOMPUTE_M.VSSX")
    $Global:AzOth = ($Global:Path+"AZUREOTHER_M.VSSX")
    $Global:AzApp = ($Global:Path+"AZUREAPPSERVICES_M.VSSX")
    $Global:AzBrick = ($Global:Path+"AZUREBLOCKCHAIN_M.VSSX")

}

Function Visio 
{

    try
        {
            $Global:application = New-Object -ComObject Visio.Application
            $Global:documents = $Global:application.Documents
            $Global:document = $Global:documents.Add("")
            $Global:pages = $Global:application.ActiveDocument.Pages
            $Global:page = $Global:pages.Item(1)
            $Global:page.Name = 'Network'

            $Global:stencil = $application.Documents.Add($AzNet)
            $Global:stenSymbol = $application.Documents.Add($AzGen)
            $Global:ComputeSymbol = $application.Documents.Add($AzCom)
            $Global:OtherSymbol = $application.Documents.Add($AzOth)
            $Global:AppSymbol = $application.Documents.Add($AzApp)
            $Global:BrickSymbol = $application.Documents.Add($AzBrick)

            $Global:Connections = $Global:stencil.Masters.Item("Connections")
            $Global:ExpressRoute = $Global:stencil.Masters.Item("ExpressRoute Circuits")
            $Global:VGW = $Global:stencil.Masters.Item("Virtual Network Gateways")
            $Global:VNET = $Global:stencil.Masters.Item("Virtual Networks")
            $Global:TRAFFIC = $Global:stencil.Masters.Item("Traffic Manager profiles")
            $Global:SymError = $Global:stenSymbol.Masters.Item("Error")
            $Global:SymInfo = $Global:stenSymbol.Masters.Item("Information")
            $Global:Subscription = $Global:stenSymbol.Masters.Item("Subscriptions")
            $Global:IconVMs = $Global:ComputeSymbol.Masters.Item("Virtual Machine")
            $Global:IconAKS = $Global:ComputeSymbol.Masters.Item("Kubernetes Services")
            $Global:IconVMSS = $Global:ComputeSymbol.Masters.Item("VM Scale Sets")
            $Global:IconLBs = $stencil.Masters.Item("Load Balancers")
            $Global:IconFWs = $Global:OtherSymbol.Masters.Item("Firewalls")
            $Global:IconPVTs = $stencil.Masters.Item("Private Link")
            $Global:IconAppGWs = $Global:OtherSymbol.Masters.Item("Application Gateways")
            $Global:IconBastions = $Global:stenSymbol.Masters.Item("Launch Portal")
            $Global:IconAPIMs = $Global:AppSymbol.Masters.Item("API Management Services")
            $Global:IconAPPs = $Global:AppSymbol.Masters.Item("App Services")
            $Global:IconBricks = $Global:BrickSymbol.Masters.Item("Azure Blockchain Service")


            #$DDoS = $stencil.Masters.Item("DDoS Protection Plans")
            #$DNS = $stencil.Masters.Item("DNS Zones")
            #$FrontDoors = $stencil.Masters.Item("Front Doors")
            #$LoadBalancers = $stencil.Masters.Item("Load Balancers")
            #$NAT = $stencil.Masters.Item("NAT")
            #$NIC = $stencil.Masters.Item("Network Interfaces")
            #$NSG = $stencil.Masters.Item("Network Security Groups")
            #$Watcher = $stencil.Masters.Item("Network Watcher")
            #$PubIP = $stencil.Masters.Item("Public IP Addresses")
        }
    catch
    {}

}



Function Network 
{

    $Global:GoldVNET = @()
    foreach($VNETTEMP1 in $AZVNETs)
        {
            foreach($VNETTEMP in $VNETTEMP1.properties.subnets.properties.ipconfigurations.id)
                {
                    $VV4 = $VNETTEMP.Split("/")
                    $VNETTEMP0 = ($VV4[0] + '/' + $VV4[1] + '/' + $VV4[2] + '/' + $VV4[3] + '/' + $VV4[4] + '/' + $VV4[5] + '/' + $VV4[6] + '/' + $VV4[7]+ '/' + $VV4[8])
                    if($VNETTEMP0 -in $AZVGWs.id)
                        {
                            $Global:GoldVNET += $VNETTEMP1
                        }
                }
        }

    $RoutsH = ($AzCons.id.count + ($AZLGWs.id).count + ($AZEXPROUTEs.id).count + (($AZVNETs.properties.virtualNetworkPeerings.properties.remoteVirtualNetwork.id | Select-Object -Unique).count))
    $RoutsW = $AZVNETs | Select-Object -Property Name, @{N="Subnets";E={$_.properties.subnets.properties.addressPrefix.count}} | Sort-Object -Property Subnets -Descending

    Start-Sleep 1

    $Ret = $page.DrawRectangle(-10,-10, (($RoutsW.Subnets[0]*1.5) +30), (($RoutsH)*3.1))

    Start-Sleep 1
    $Global:Alt = 2


    foreach($GTW in $AZLGWs)
        {
                $vvnet = $page.Drop($TRAFFIC, 4.5, $Global:Alt) 
                if($GTW.properties.provisioningState -ne 'Succeeded')
                {
                    $ErrorLC = $page.Drop($SymError, 3.7, $Global:Alt)
                    $ErrorLC.Text = 'Failed'
                }
        
                $Con1 = $AZCONs | where {$_.properties.localNetworkGateway2.id -eq $GTW.id}
                Diagram $Con1
                if(!$Con1 -and $GTW.properties.provisioningState -eq 'Succeeded')
                {
                    $InfoLC = $page.Drop($SymInfo, 3.1, $Global:Alt)
                    $InfoLC.Text = 'No Connection Found'            
                }
                $Global:Alt = $Global:Alt + 2
                $Name = $GTW.name
                $IP = $GTW.properties.gatewayIpAddress
                $vvnet.Text = ([string]$Name + "`n" + [string]$IP)        
        }


####################################################################### ERS #####################################################################################

    $Global:Alt = $Global:Alt + 4 
    Foreach($ERs in $AZEXPROUTEs)
        {
                $vvnet = $page.Drop($ExpressRoute, 4.5, $Global:Alt) 
                if($GTW.properties.provisioningState -ne 'Succeeded')
                {
                    $ErrorLC = $page.Drop($SymError, 3.7, $Global:Alt)
                    $ErrorLC.Text = 'Failed'
                }
        
                $Con1 = $AZCONs | where {$_.properties.peer.id -eq $ERs.id}
                Diagram $Con1
                if(!$Con1 -and $ERs.properties.circuitProvisioningState -eq 'Enabled')
                {
                    $InfoLC = $page.Drop($SymInfo, 3.1, $Global:Alt)
                    $InfoLC.Text = 'No Connection Found'
                }
                $Global:Alt = $Global:Alt + 2
                $Name = $ERs.name
                $vvnet.Text = [string]$Name  


        $OnPrem = $page.DrawRectangle(-2, 1, 4, $Global:Alt)
        $OnPrem.Text = 'OnPrem Environment'

        }

}


Function Diagram 
{
Param($Con1)
foreach ($Con2 in $Con1)
        {
            $Global:vnetLoc = 10.3
            $VGT = $AZVGWs | where {$_.id -eq $Con2.properties.virtualNetworkGateway1.id}
            $Conn = $page.Drop($Connections, 6, $Global:Alt)
            $Name2 = $Con2.Name
            $Conn.Text = [string]$Name2
            $vvnet.AutoConnect($Conn,0)

            $vpngt = $page.Drop($VGW, 8, $Global:Alt)
            $vpngt.Text = [string]$VGT.Name
            $Conn.AutoConnect($vpngt,0)

            foreach($AZVNETs2 in $AZVNETs)
            {
                foreach($VNETTEMP in $AZVNETs2.properties.subnets.properties.ipconfigurations.id)
                {
                    $VV4 = $VNETTEMP.Split("/")
                    $VNETTEMP1 = ($VV4[0] + '/' + $VV4[1] + '/' + $VV4[2] + '/' + $VV4[3] + '/' + $VV4[4] + '/' + $VV4[5] + '/' + $VV4[6] + '/' + $VV4[7]+ '/' + $VV4[8])
                    if($VNETTEMP1 -eq $VGT.id)
                    {
                        $VNET2 = $AZVNETs2

                        $VNETsD = $page.Shapes | where {$_.name -like 'Virtual Networks*'}
                        $Global:Alt0 = $Global:Alt
                        $vali = 0
                        if(($VNET2.name + "`n" + $VNET2.properties.addressSpace.addressPrefixes) -notin $VNETsD.text)
                            {
                                $vali = 1
                                $vpnnet = $page.Drop($VNET, 10, $Global:Alt)
                                if($VNET2.properties.addressSpace.addressPrefixes.count -ge 10){$AddSpace = ($VNET2.properties.addressSpace.addressPrefixes | Select-Object -First 10)+ "`n" +'...'}Else{$AddSpace = $VNET2.properties.addressSpace.addressPrefixes}
                                $vpnnet.Text = ([string]$VNET2.Name + "`n" + $AddSpace)
                                $vpngt.AutoConnect($vpnnet,0)
                                                    
                                if($VNET2.properties.virtualNetworkPeerings.properties.remoteVirtualNetwork.id)
                                    {
                                        $PeerCount = ($VNET2.properties.virtualNetworkPeerings.properties.remoteVirtualNetwork.id.count + 10.3)
                                        $Global:vnetLoc1 = $Global:Alt
                                       
                                        if($VNET2.properties.subnets.properties.addressPrefix.count -gt 5)
                                            {
                                                $Global:vnetLoc1 = $Global:vnetLoc1 + 5
                                            }
                                            else
                                            {
                                                $Global:vnetLoc1 = $Global:vnetLoc1 + 3
                                            }

                                        Foreach ($Peer in $VNET2.properties.virtualNetworkPeerings)
                                            {
                                                $VNETSUB = $AZVNETs | where {$_.id -eq $Peer.properties.remoteVirtualNetwork.id}                                                
                                                
                                                if(($VNETSUB.name + "`n" + $VNETSUB.properties.addressSpace.addressPrefixes) -in $VNETsD.text)
                                                    {
                                                        $VNETDID = $VNETsD | where {$_.Text -eq ($VNETSUB.name + "`n" + $VNETSUB.properties.addressSpace.addressPrefixes)}
                                                        $vpnnet.AutoConnect($VNETDID,0)
                                                        $Conn2 = $page.Shapes | where {$_.name -like 'Dynamic connector*'} | select-object -Last 1
                                                        $Conn2.text = $Peer.name
                                                        $Conn2.Cells('BeginArrow')=4
                                                        $Conn2.Cells('EndArrow')=4
                                                    }
                                                elseif($Peer.properties.remoteVirtualNetwork.id -notin $GoldVNET.id -and $Peer.properties.peeringState -ne 'Disconnected')
                                                {                                                      
                                                    $Global:sizeL =  $VNETSUB.properties.subnets.properties.addressPrefix.count   
                                                                                                                                    
                                                    $Global:vnetLoc = 12
                                                     
                                                    $netpeer = $page.Drop($VNET, $Global:vnetLoc, $Global:vnetLoc1)                                            
                                                    $netpeer.AutoConnect($vpnnet,0)
                                                    $Conn1 = $page.Shapes | where {$_.name -like 'Dynamic connector*'} | select-object -Last 1
                                                    $Conn1.text = $Peer.name
                                                    $Conn1.Cells('BeginArrow')=4
                                                    $Conn1.Cells('EndArrow')=4                        
                                            
                                                    $netpeer.Text = ($VNETSUB.name + "`n" + $VNETSUB.properties.addressSpace.addressPrefixes)                                            
                                                
                                                    if ($Global:sizeL -gt 5)
                                                        {
                                                            $Global:sizeL = $Global:sizeL / 2
                                                            $Global:sizeL = [math]::ceiling($Global:sizeL)
                                                            $Global:sizeC = $Global:sizeL
                                                            $Global:sizeL = ($Global:sizeL*1.5)+(12+0.7)
                                                            $vnetbox = $page.DrawRectangle((12+0.5), ($Global:vnetLoc1 - 0.5), $Global:sizeL, ($Global:vnetLoc1 + 3.3))                                                                                                                      

                                                            $SubIcon = $page.Drop($Subscription, ($Global:sizeL), ($Global:vnetLoc1-0.6))
                                                            $SubName = $Subscriptions | Where {$_.id -eq $VNETSUB.subscriptionId}
                                                            $SubIcon.Text = $SubName.name
                                                
                                                            $Global:subloc0 = (12+0.6)
                                                            $Global:SubC = 0
                                                            foreach($Global:Sub in $VNETSUB.properties.subnets)
                                                                {
                                                                    if ($Global:SubC -eq $Global:sizeC) 
                                                                        {
                                                                            $Global:vnetLoc1 = $Global:vnetLoc1 + 1.7                                         
                                                                            $Global:subloc0 = (12+0.6)
                                                                            $Global:SubC = 0
                                                                        }
                                                                    $vsubnetbox = $page.DrawRectangle($Global:subloc0, ($Global:vnetLoc1 - 0.3), ($Global:subloc0 + 1.5), ($Global:vnetLoc1 + 1.3))
                                                                    $vsubnetbox.Text = ("`n" + "`n" + "`n" + "`n" + "`n" + [string]$sub.Name + "`n" + [string]$sub.properties.addressPrefix)
                                                                    
                                                                    ProcType $Global:sub $Global:subloc0 $Global:vnetLoc1
                                                                    
                                                                    $Global:subloc0 = $Global:subloc0 + 1.5
                                                                    $Global:SubC ++

                                                                }                                                                
                                                        }
                                                    else
                                                        {
                                                            $Global:sizeL = ($Global:sizeL*1.5)+(12+0.7)
                                                            $vnetbox = $page.DrawRectangle((12+0.5), ($Global:vnetLoc1 - 0.5), $Global:sizeL, ($Global:vnetLoc1 + 1.6))

                                                            $SubIcon = $page.Drop($Subscription, ($Global:sizeL), ($Global:vnetLoc1-0.6))
                                                            $SubName = $Subscriptions | Where {$_.id -eq $VNETSUB.subscriptionId}
                                                            $SubIcon.Text = $SubName.name

                                                            $Global:subloc0 = (12+0.6)
                                                            foreach($Global:sub in $VNETSUB.properties.subnets)
                                                                {
                                                                    $vsubnetbox = $page.DrawRectangle($Global:subloc0, ($Global:vnetLoc1 - 0.3), ($Global:subloc0 + 1.5), ($Global:vnetLoc1 + 1.3))
                                                                    $vsubnetbox.Text = ("`n" + "`n" + "`n" + "`n" + "`n" +[string]$sub.Name + "`n" + [string]$sub.properties.addressPrefix)

                                                                    ProcType $Global:sub $Global:subloc0 $Global:vnetLoc1

                                                                    $Global:subloc0 = $Global:subloc0 + 1.5
                                                                }
                                                        }
                                                    if($Global:sizeL -gt 5)
                                                        {
                                                            $Global:Alt = $Global:Alt + 2.5
                                                        }
                                                    else
                                                        {
                                                            $Global:Alt = $Global:Alt + 1.5
                                                        }
                                                    $Global:vnetLoc1 = $Global:vnetLoc1 + 3                                            
                                                }                          
                                            }
                                    }
                            }
                        else
                            {
                                $VNETDID = $VNETsD | where {$_.Text -eq ($VNET2.name + "`n" + $VNET2.properties.addressSpace.addressPrefixes)}
                                $vpngt.AutoConnect($VNETDID,0)
                                $Conn2.Cells('BeginArrow')=4
                                $Conn2.Cells('EndArrow')=4
                            }
                        if ($vali -eq 1)
                            {
                                $Global:sizeL =  $VNET2.properties.subnets.properties.addressPrefix.count
                                if ($Global:sizeL -gt 5)
                                {
                                    $Global:sizeL = $Global:sizeL / 2
                                    $Global:sizeL = [math]::ceiling($Global:sizeL)
                                    $Global:sizeC = $Global:sizeL
                                    $Global:sizeL = ($Global:sizeL*1.5)+($Global:vnetLoc + 0.7)
                                    $vnetbox = $page.DrawRectangle(($Global:vnetLoc+0.5), ($Global:Alt0 - 0.5), $Global:sizeL, ($Global:Alt0 + 3.3))

                                    $SubIcon = $page.Drop($Subscription, ($Global:sizeL), ($Global:Alt0-0.6))
                                    $SubName = $Subscriptions | Where {$_.id -eq $VNET2.subscriptionId}
                                    $SubIcon.Text = $SubName.name

                                    $Global:subloc = ($Global:vnetLoc+0.6)
                                    $Global:SubC = 0
                                    foreach($Global:Sub in $VNET2.properties.subnets)
                                    {
                                        if ($Global:SubC -eq $Global:sizeC) 
                                        {
                                            $Global:Alt0 = $Global:Alt0 + 1.7
                                            $Global:subloc = ($Global:vnetLoc+0.6)
                                            $Global:SubC = 0
                                        }
                                        $vsubnetbox = $page.DrawRectangle($Global:subloc, ($Global:Alt0 - 0.3), ($Global:subloc + 1.5), ($Global:Alt0 + 1.3))
                                        $vsubnetbox.Text = ("`n" + "`n" + "`n" + "`n" + "`n" + [string]$sub.Name + "`n" + [string]$sub.properties.addressPrefix)

                                        ProcType $Global:sub $Global:subloc $Global:Alt0

                                        $Global:subloc = $Global:subloc + 1.5
                                        $Global:SubC ++
                                    }
                                }
                            else
                                {
                                    $Global:sizeL = ($Global:sizeL*1.5)+($Global:vnetLoc + 0.7)
                                    $vnetbox = $page.DrawRectangle(($Global:vnetLoc+0.5), ($Global:Alt0 - 0.5), $Global:sizeL, ($Global:Alt0 + 1.6))

                                    $SubIcon = $page.Drop($Subscription, ($Global:sizeL), ($Global:Alt0-0.6))
                                    $SubName = $Subscriptions | Where {$_.id -eq $VNET2.subscriptionId}
                                    $SubIcon.Text = $SubName.name

                                    $Global:subloc = ($Global:vnetLoc+0.6)
                                    foreach($Global:Sub in $VNET2.properties.subnets)
                                    {
                                        $vsubnetbox = $page.DrawRectangle($Global:subloc, ($Global:Alt0 - 0.3), ($Global:subloc + 1.5), ($Global:Alt0 + 1.3))
                                        $vsubnetbox.Text = ("`n" + "`n" + "`n" + "`n" + "`n" +[string]$sub.Name + "`n" + [string]$sub.properties.addressPrefix)

                                        ProcType $Global:sub $Global:subloc $Global:Alt0

                                        $Global:subloc = $Global:subloc + 1.5
                                    }
                                }
                            if($Global:sizeL -gt 5)
                                {
                                    $Global:Alt = $Global:Alt + 2.5
                                }
                            else
                                {
                                    $Global:Alt = $Global:Alt + 1.5
                                }                                              
                        }
                    }
                }
            }
            if($Con1.count -gt 1)
            {
               $Global:Alt = $Global:Alt + 1#3
            }
        }

}



Function ProcType 
{
Param($sub,$subloc,$Alt0)
    $Types = @()
    Remove-Variable temp -ErrorAction SilentlyContinue
    Remove-Variable TrueTemp -ErrorAction SilentlyContinue
    foreach($type in $sub.properties.ipconfigurations.id)
        {
            if($type.Split("/")[7] -eq 'virtualMachineScaleSets' -and $type.Split("/")[8] -like 'aks-*')
                {
                    $Types += 'AKS'
                }
            else
                {
                    $Types += $type.Split("/")[7]
                }
        }
    $temp = $types | Group-Object | Sort-Object -Property Count -Descending
    if ([string]::IsNullOrEmpty($temp))
        {
            if($sub.properties.resourceNavigationLinks.properties.linkedResourceType -eq 'Microsoft.ApiManagement/service')
            {
                $TrueTemp = 'APIM'
            }
            if($sub.properties.serviceAssociationLinks.properties.link)
            {
                if($sub.properties.serviceAssociationLinks.properties.link.split("/")[6] -eq 'Microsoft.Web')
                    {
                        $TrueTemp = 'APP Service'
                    }
            }
            if($sub.properties.applicationGatewayIPConfigurations.id)
            {
                if($sub.properties.applicationGatewayIPConfigurations.id.split("/")[7] -eq 'applicationGateways')
                    {
                        $TrueTemp = 'applicationGateways'
                    }
            }
            if($sub.properties.delegations.properties.serviceName)
            {
                if($sub.properties.delegations.properties.serviceName.split("/")[0] -eq 'Microsoft.Web')
                    {
                        $TrueTemp = 'APP Service'
                    }
            }
            if($sub.properties.delegations.properties.serviceName)
            {
                if($sub.properties.delegations.properties.serviceName.split("/")[0] -eq 'Microsoft.Databricks')
                    {
                        $TrueTemp = 'DataBricks'
                    }
            }                                                                                                                
        }
    else
        {
            $TrueTemp = $temp[0].name
        }
    switch ($TrueTemp)
        {
            'networkInterfaces' {
                                $SubIcon = $page.Drop($IconVMs, ($subloc+0.75), ($Alt0+0.8))
                                $SubIcon.Text = ([string]$temp[0].Count + ' VMs')                                                                        
                                }
            'AKS' {                                                    
                                $SubIcon = $page.Drop($IconAKS, ($subloc+0.75), ($Alt0+0.8))
                                $SubIcon.Text = ([string]$temp[0].Count + ' AKS Nodes')                                                                        
                                }
            'virtualMachineScaleSets' {                                                    
                                $SubIcon = $page.Drop($IconVMSS, ($subloc+0.75), ($Alt0+0.8))
                                $SubIcon.Text = ([string]$temp[0].Count + ' VMSS VMs')                                                                        
                                } 
            'loadBalancers' {                                                    
                                $SubIcon = $page.Drop($IconLBs, ($subloc+0.75), ($Alt0+0.8))
                                $SubIcon.Text = ([string]$temp[0].Count + ' Load Balancers')                                                                        
                                } 
            'virtualNetworkGateways' {                                                    
                                $SubIcon = $page.Drop($VGW, ($subloc+0.75), ($Alt0+0.8))
                                $SubIcon.Text = ([string]$temp[0].Count + ' Virtual Network Gateway')                                                                        
                                } 
            'azureFirewalls' {                                                    
                                $SubIcon = $page.Drop($IconFWs, ($subloc+0.75), ($Alt0+0.8))
                                $SubIcon.Text = ([string]$temp[0].Count + ' Azure Firewalls')                                                                        
                                } 
            'privateLinkServices' {                                                    
                                $SubIcon = $page.Drop($IconPVTs, ($subloc+0.75), ($Alt0+0.8))
                                $SubIcon.Text = ([string]$temp[0].Count + ' Private Links')                                                                        
                                } 
            'applicationGateways' {                                                    
                                $SubIcon = $page.Drop($IconAppGWs, ($subloc+0.75), ($Alt0+0.8))
                                $SubIcon.Text = ([string]$temp[0].Count + ' Application Gateway')                                                                        
                                }                     
            'bastionHosts' {                                                    
                                $SubIcon = $page.Drop($IconBastions, ($subloc+0.75), ($Alt0+0.8))
                                $SubIcon.Text = ([string]$temp[0].Count + ' Bastion Host')                                                                        
                                } 
            'APIM' {                                                    
                                $SubIcon = $page.Drop($IconAPIMs, ($subloc+0.75), ($Alt0+0.8))
                                $SubIcon.Text = 'APIM'                                                                   
                                }
            'APP Service' {                                                    
                                $SubIcon = $page.Drop($IconAPPs, ($subloc+0.75), ($Alt0+0.8))
                                $SubIcon.Text = 'APP Service'                                                                    
                                }
            'DataBricks' {                                                    
                                $SubIcon = $page.Drop($IconBricks, ($subloc+0.75), ($Alt0+0.8))
                                $SubIcon.Text = 'Azure Databricks'                                                                    
                                }                                                                                                        
            '' {}
            default {}
        }
}


Function VMs {

$pages.Add("")
$page = $pages.Item(2)
$page.Name = 'Virtual Machines'

($AZVMs | Select-Object -Property resourceGroup -Unique).count

$res = $resources | where {$_.type -notin ('microsoft.compute/disks','microsoft.compute/virtualmachines/extensions','microsoft.network/networkinterfaces','microsoft.compute/availabilitysets','microsoft.web/sites/slots','microsoft.portal/dashboards')}

$res | Group-Object -Property resourceGroup -NoElement | Sort-Object -Property Count -Descending | where {$_.Count -ge 10}

$res | where {$_.resourceGroup -eq 'b2c-sso'} | group -Property type | Select-Object 'Count','Name' | Sort-Object -Property 'Count' -Descending

}



Variables0
Visio
Network
#VMs

$Global:document.SaveAs($DFile) | Out-Null
$Global:application.Quit() 