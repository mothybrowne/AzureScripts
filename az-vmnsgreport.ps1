#Set Arrays
$azvmarray = @()
$info = @()
$nsginfo = @()
$allobj = @()
$finalserverreport = @()
#Set other variables
$azsubs = get-azsubscription
$date = get-date -format "ddMMyyHHmm"
$mydocuments = [environment]::getfolderpath("mydocuments")
$documentName = ($mydocuments + '\NSGReport' + $date + '.xlsx')
Connect-AzAccount
#Run through each Subscription
ForEach ($vsub in $azsubs)
    {
    Select-AzSubscription $vsub.SubscriptionID
    Write-Host "Working on Subscription:" $vsub.Name
    $azvms = @()
    $aznsgs = @()
    $vms = get-azvm -Status
    $nics = get-aznetworkinterface | ?{ $_.VirtualMachine -NE $null}
    
    
#Build Array of Each VM with IP and NIC 
    foreach($nic in $nics)
        {
        #$info = "" | Select VmName, ResourceGroupName, HostName, NIC, IpAddress, PublicIP, PowerStatus, Location, OsType, VmSize
        $vm = $vms | ? -Property Id -eq $nic.VirtualMachine.id
        if ($nic.IpConfigurations.PublicIpAddress.Id)
        {
            $extnicname = $nic.IpConfigurations.PublicIpAddress.Id.Split('/') | select -last 1
            $PublicIP = (Get-AzPublicIpAddress -Name $extnicname).IpAddress
        } else 
        {
            $PublicIP = 'NULL'
        }
        $obvm = $null
        $obvm = [PSCustomObject] @{
        VMName = $vm.Name
        HostName = $vm.OSProfile.ComputerName
        ResourceGroupName = $vm.ResourceGroupName
        PrivateIP = $nic.IpConfigurations.PrivateIpAddress
        PublicIP = $PublicIP
        NICName = $nic.Name
        PowerStatus = $vm.PowerState
        OsType = $vm.StorageProfile.OSDisk.OSType
        VmSize = $vm.hardwareprofile.vmsize
        Location = $vm.Location
        }
        $azvms+=$obvm
        }
    $finalserverreport += $azvmsfor
    Write-Host "Servers to Audit:" $azvms.count
#Associcate each NIC with it's NSG
    for ( $ix = 0; $ix -lt $azvms.count; $ix++ )
        {
        Write-Host $azvms[$ix].VMName $azvms[$ix].NICName
        $ansg = Get-AzNetworkInterface -name $azvms[$ix].NICName
        if (!$ansg.NetworkSecurityGroup.Id)
        {
            $ansgname = "Internal NSG"
        } else
        {
            $ansgname = $ansg.NetworkSecurityGroup.Id | %{ $_.Split('/')[-1];}
        }
        $nsginfo = "" | Select VmName, NSGName
        $nsginfo.VmName = $azvms[$ix].VmName
        $nsginfo.NSGName = $ansgname
        $aznsgs+=$nsginfo
        }
    Write-Host "NSG's to Audit:" $aznsgs.count
#Audit the Network Security Groups
    $nsgs = Get-AzNetworkSecurityGroup
    foreach ($nsg in $nsgs)
        {
        $subnet = $nsg.Subnets
        if ($subnet.Count -gt 0) {
            $snarrayid = $subnet.id.Split('/').count - 1
            $vnarrayid = $subnet.id.Split('/').count - 3
            $snarray = $subnet.id.Split('/')
            $vnetname = $snarray[$vnarrayid]
            $subnetname = $snarray[$snarrayid]
        } else {
            $vnetname = '(NOT ASSIGNED)'
            $subnetname = '(NOT ASSIGNED)'
        }
        $azvmname = $aznsgs | where { $_.NSGName -eq $nsg.Name }
        #write-host $azvmname.VmName
        $rules = $nsg.SecurityRules
        foreach ($rule in $rules)
        {
            $obnsg = $null
            $obnsg = [PSCustomObject] @{
            NSGName = $nsg.Name
            ServerName = $azvmname.VmName
            ResourceGroup = $nsg.ResourceGroupName
            Location = $nsg.Location
            RuleName = $rule.Name
            Priority = $rule.Priority.ToString() 
            Direction = $rule.Direction
            Access = $rule.Access
            SourceAddress = $rule.SourceAddressPrefix
            SourcePort = $rule.SourcePortRange
            DestinationAddress = $rule.DestinationAddressPrefix 
            DestinationPort = $rule.DestinationPortRange
            Description = $rule.Description
            VNet = $vnetname
            Subnet = $subnetname
            }
            $allobj += $obnsg
        }
    }
    
    
    
}
#Strip Object lists ready for export
$jsonconv = $allobj | ConvertTo-Json -depth 1 
$finalreport = $jsonconv | convertfrom-json
#Export to Excel
$finalserverreport | Export-Excel -WorksheetName "Servers" $documentName -AutoSize -FreezeTopRow -AutoFilter
$finalreport | Export-Excel -WorksheetName "NSG Rules" $documentName -AutoSize -FreezeTopRow -AutoFilter
