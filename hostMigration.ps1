<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2016 v5.2.120
	 Created on:   	01/18/2016 10:40 AM
	 Created by:    Charles Welton (WeltonWare)	 
	 Organization: 	MTS
	 Filename:     	hostMigration.ps1
	===========================================================================
	.DESCRIPTION
		A description of the file.
#>


Write-Host "Ensuring ALL VI server connections are disconnected before starting..."
if ($global:DefaultVIServers)
{
	Disconnect-VIServer -Server $global:DefaultVIServers -Force -Confirm:$false | Out-Null
}

#Configure initial window size and title
$pshost                   = get-host
$pswindow                 = $pshost.ui.rawui
$pswindow.windowtitle     = "WeltonWare VMWare Host Migration"
$pswindow.foregroundcolor = "White"
$pswindow.backgroundcolor = "Black"
$newsize                  = $pswindow.windowsize
$newsize.height           = 30
$newsize.width            = 100
$pswindow.windowsize      = $newsize

#Suppress all "warning" messages
$warningPreference        = "SilentlyContinue"

$dateFormat               = Get-Date -Format yyyy_d_M
$basicClr                 = "white"
$stdMsgClr                = "DarkGray"
$actionClr                = "blue"
$errClr                   = "red"
$successClr               = "green"
$waitMsgClr               = "magenta"
$powerStateClr            = "DarkRed"
$WWLabelBackClr           = "yellow"
$WWLabelForeClr           = "black"
$menuItemClr              = "gray"

[System.Console]::Clear()

#Check for registered VMware snap-ins and register them if they are not already registered
$snapins = Get-PSSnapin -Registered -Name VMware*
foreach ($snapin in $snapins)
{
	if (!(Get-PSSnapin -name $snapin.name -erroraction silentlycontinue))
	{
		Add-PSSnapin $snapin.name
	}
}

function promptForFile()
{
	[System.Console]::Clear()
	$dataFile = "hostMigration_data.xlsx"
	
	Write-Host "Please load file named `"${dataFile}`"..." -ForegroundColor "${stdMsgClr}"
	sleep 2
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
	$fd                  = New-Object system.windows.forms.openfiledialog
	$fd.InitialDirectory = "${dataFile}"
	$fd.MultiSelect      = $true
	$fd.showdialog()
	$fd.filenames
	
	if ($fd.filenames -and $fd.filenames -match "${dataFile}")
	{
		
		$xl         = New-Object -COM "Excel.Application"
		$xl.Visible = $false
		$wb         = $xl.Workbooks.Open($fd.filenames)
		$ws         = $wb.Sheets.Item(1)
		
		Write-Host "Loading data from file..." -ForegroundColor "${actionClr}"
		for ($i = 1; $i -le 17; $i++)
		{
			if ($i -eq 11)
			{
				$script:sourceVI = $ws.Cells.Item($i, 2).text
			}
			
			if ($i -eq 12)
			{
				$script:destVI = $ws.Cells.Item($i, 2).text
			}
			
			if ($i -eq 13)
			{
				$script:sourceHost = $ws.Cells.Item($i, 2).text
			}
			
			if ($i -eq 14)
			{
				$script:credentials = $ws.Cells.Item($i, 2).text -split '\\'
				$script:vCenterUser = $credentials[0]
				$script:vCenterPass = $credentials[1]
			}
			
			if ($i -eq 15)
			{
				$script:credentials = $ws.Cells.Item($i, 2).text -split '\\'
				$script:hostUser    = $credentials[0]
				$script:hostPass    = $credentials[1]
			}
			
			if ($i -eq 16)
			{
				$script:sourceDC = $ws.Cells.Item($i, 2).text
			}
			
			if ($i -eq 17)
			{
				$script:sourceCluster = $ws.Cells.Item($i, 2).text
			}
			
		}
		
		$wb.Close()
		$xl.Quit()
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
		
		# ------- VLAN Text File Data -------
		$script:VLANinfo = "c:\NetworkVLANInformation_${SourceHost}_${dateFormat}.txt"
		
		# ------- Global Templates File -------
		$global:templatesList = "c:\template_list_${SourceHost}_${dateFormat}.txt"
	}
	else
	{
		Write-Host "No data file!" -ForegroundColor "${errClr}"
		Write-Host " "
		exit
	}
}

function connectVI($VI)
{
	Write-Host "Connecting to ${VI}..." -ForegroundColor "${actionClr}"
	Connect-VIServer –Server $VI -User $vCenterUser -Password $vCenterPass -ea SilentlyContinue | Out-Null
	if ($? -eq "True")
	{
		Write-Host "Successfully logged in!" -ForegroundColor "${successClr}"
		Write-Host " "
	}
	else
	{
		Write-Host "Failed to login!" -ForegroundColor "${errClr}"
		Write-Host " "
		exit
	}
}

function disconnectVI()
{
	Write-Host "Disconnecting all server connections..." -ForegroundColor "${actionClr}"
	Write-Host " "
	
	Disconnect-VIServer -Server * -Force -confirm:$false | Out-Null
}

function Get-dvPgFreePort
{
	param (
		[CmdletBinding()]
		[string]$PortGroup,
		[int]$Number = 1
	)
	
	$nicTypes = "VirtualE1000", "VirtualE1000e", "VirtualPCNet32",
	"VirtualVmxnet", "VirtualVmxnet2", "VirtualVmxnet3"
	$ports = @{ }
	
	$pg = Get-VirtualPortGroup -Distributed -Name $PortGroup
	$pg.ExtensionData.PortKeys | %{ $ports.Add($_, $pg.Name) }
	
	if ($pg.ExtensionData.vm.count -gt 0)
	{
		Get-View $pg.ExtensionData.Vm | %{
			$nic = $_.Config.Hardware.Device |
			where {
				$nicTypes -contains $_.GetType().Name -and
				$_.Backing.GetType().Name -match "Distributed"
			}
			$nic | %{ $ports.Remove($_.Backing.Port.PortKey) }
		}
	}
	
	if ($Number -gt $ports.Keys.Count)
	{
		$Number = $ports.Keys.Count
	}
	($ports.Keys | Sort-Object)[0..($Number - 1)]
}

function Get-FolderByPath
{
	param (
		[CmdletBinding()]
		[parameter(Mandatory = $true)]
		[System.String[]]${Path},
		[char]${Separator} = '/'
	)
	
	process
	{
		if ((Get-PowerCLIConfiguration).DefaultVIServerMode -eq "Multiple")
		{
			$vcs = $defaultVIServers
		}
		else
		{
			$vcs = $defaultVIServers[0]
		}
		
		foreach ($vc in $vcs)
		{
			foreach ($strPath in $Path)
			{
				$root = Get-Folder -Name Datacenters -Server $vc
				$strPath.Split($Separator) | %{
					$root = Get-Inventory -Name $_ -Location $root -Server $vc -NoRecursion
		
					if ((Get-Inventory -Location $root -NoRecursion | Select -ExpandProperty Name) -contains "vm")
					{
						$root = Get-Inventory -Name "vm" -Location $root -Server $vc -NoRecursion
						
					}
				}
				
				$root | where { $_ -is [VMware.VimAutomation.ViCore.Impl.V1.Inventory.FolderImpl] } | %{
					Get-Folder -Name $_.Name -Location $root.Parent -NoRecursion -Server $vc
				}
			}
		}
	}
}

function HostMigration()
{
	if (promptForFile)
	{
		if (!$sourceHost -or !$hostUser -or !$hostPass)
		{
			Write-Host "Host information is missing from the data file!" -ForegroundColor "${errClr}"
			Write-Host " "
			exit
		}
		#Change to multi-mode vcenter management
		#Set-PowerCLIConfiguration -DefaultVIServerMode Multiple -Confirm:$false
		Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false | Out-Null
		
		function getVM()
		{
			Get-VM | %{ $_.Name }
		}
		
		function getVLANInfo($MyHOST)
		{
			Write-Host "Getting VLAN information for ${MyHOST}..." -ForegroundColor "${actionClr}"
			Write-Host "This is going to take some time, please be patient!" -Foreground "${waitMsgClr}"
			Write-Host " "
			if (Test-Path -Path ${VLANinfo})
			{
				Remove-Item -Force ${VLANinfo}
			}
			
			$filterVMName = Get-VMHost ${MyHOST} | select -Unique | Get-VM | Where-Object { $_.Host.Name -eq ${MyHOST} } | %{ $_.Name }
			$filterArray  = @($filterVMName)
			$filterArray  = $filterArray | Where-Object { $_ } | Select -Unique
			$filterCount  = [int]$filterArray.Count
			$loopCount    = 1
			
			if ($filterCount -gt 0)
			{
				$script:VDS = GetVDSwitch $MyHOST
			
				foreach ($f in $filterArray)
				{
					$vmPortKey = ((get-vmhost -name "${MyHOST}" | get-vm -name "${f}").NetworkAdapters).extensiondata.backing.port.portkey
					Get-NetworkAdapter -VM ${f} | %{
						$NetAdapterName = $_.Name
			
						$getPortGrp = (get-vm "${f}" | get-view).network.value
			
						if ($getPortGrp)
						{
							$NetworkName = $_.NetworkName
						} else {
							$NetworkName = " "
						}
			
						$Type = $_.Type
						$MAC  = $_.MacAddress
					}
					echo "${VDS}	${f}	${NetAdapterName}	${NetworkName}	${Type}	${MAC}	${vmPortKey}" | Add-Content ${VLANinfo}
				}
				
			convertTemplates_to_VMs
			DisconnectHost $MyHOST
			}
		}
		
		function DisconnectHost($MyHOST)
		{
			Write-Host "Attempting to disconnect ${MyHOST} from cluster ${sourceVI}..." -ForegroundColor "${actionClr}"
			Write-Host " "
			Set-VMHost -VMHost $MyHOST -State "Disconnected" | Out-Null
			
			if ($? -eq "True")
			{
				Write-Host "Disconnecting of host was successful!  Continuing..." -ForegroundColor "${successClr}"
				write-Host " "
			}
			else
			{
				Write-Host "Disconnecting of host failed for some reason!  Exiting." -ForegroundColor "${errClr}"
				write-Host " "
				exit
			}
			
			#RemoveHost $MyHOST $sourceVI
		}
		
		function RemoveHost($MyHOST, $cluster)
		{
			Write-Host "Attempting to remove ${MyHOST} from cluster ${cluster}..." -ForegroundColor "${actionClr}"
			Write-Host " "
			Remove-VMHost -VMHost $MyHOST -Confirm:$false
		}
		
		function AddHost($MyDC, $MyHOST, $VDS)
		{
			$script:pnicArray = @()
			$script:pnicArray = (get-vmhostnetworkadapter -vmhost $MyHOST -physical -virtualswitch $VDS).Name		
			$script:getUplink = ((get-vdport -vdswitch $VDS -uplink -connectedonly -activeonly).Portgroup).Name | select -First 1
			
			disconnectVI
			connectVI $destVI
			
			if ($? -eq "True")
			{
				Write-Host "Attempting to add ${MyHOST} to cluster ${destVI}..." -ForegroundColor "${actionClr}"
				Write-Host " "
				Add-VMHost $MyHOST -Location $sourceCluster -Force -User $hostUser -Password $hostPass | Out-Null
				
				if ($? -eq "True")
				{
					Write-Host "Adding of host was successful!  Continuing..." -ForegroundColor "${successClr}"
					write-Host " "
					convertVMs_to_Templates
					migrateNetworking $MyDC $VDS $MyHOST
					add_VM_NICs
				}
				else
				{
					Write-Host "Adding of host failed for some reason!  Exiting." -ForegroundColor "${errClr}"
					write-Host " "
					exit
				}
			}
		}
		
		function GetVDSwitch($MyHOST)
		{
			Get-VMHost -Name $MyHOST | Get-VDSWitch | %{ $_.Name }
		}
		
		function migrateNetworking($dc, $vds, $vmhost)
		{
			$script:failedVMList   = "c:\vm_failed_list_${vmhost}_${dateFormat}.txt"
			$script:completeVMList = "c:\vm_complete_list_${vmhost}_${dateFormat}.txt"
			
			if (Test-Path -Path "${failedVMList}")
			{
				Remove-Item $failedVMList
			}
			
			if (Test-Path -Path "${completeVMList}")
			{
				Remove-Item $completeVMList
			}
			
			Write-Host "Collecting spec data..." -ForegroundColor "${actionClr}"
			
			# ------- QueryCompatibleHostForExistingDvs -------
			
			$container       = New-Object VMware.Vim.ManagedObjectReference
			$container.type  = "Datacenter"
			$container.Value = (get-datacenter $dc | get-view).moref.value
			$containerVal    = $container.Value
			
			$dvs             = New-Object VMware.Vim.ManagedObjectReference
			$dvs.type        = "VmwareDistributedVirtualSwitch"
			$dvs.Value       = (get-vdswitch $vds | get-view).moref.value
			$dvsVal          = $dvs.Value
			
			$_this           = Get-View -Id 'DistributedVirtualSwitchManager-DVSManager'
			$_this.QueryCompatibleHostForExistingDvs($container, $true, $dvs)
			
			
			# ------- FetchDVPortKeys -------
			
			$criteria            = New-Object VMware.Vim.DistributedVirtualSwitchPortCriteria
			$criteria.connected  = $true
			$criteria.uplinkPort = $false
			
			
			$_this = Get-View -Id "VmwareDistributedVirtualSwitch-${dvsVal}"
			$_this.FetchDVPortKeys($criteria)
			
			
			# ------- FetchDVPorts -------
			
			$criteria            = New-Object VMware.Vim.DistributedVirtualSwitchPortCriteria
			$criteria.uplinkPort = $true
			
			$_this = Get-View -Id "VmwareDistributedVirtualSwitch-${dvsVal}"
			$_this.FetchDVPorts($criteria)
			
			
			# ------- ReconfigureDvs_Task -------
			
			$dvSw                    = Get-View -Id "VmwareDistributedVirtualSwitch-${dvsVal}"
			
			$spec                    = New-Object VMware.Vim.DVSConfigSpec
			$spec.configVersion      = $dvSw.Config.ConfigVersion
			$spec.host               = New-Object VMware.Vim.DistributedVirtualSwitchHostMemberConfigSpec[] (1)
			$spec.host[0]            = New-Object VMware.Vim.DistributedVirtualSwitchHostMemberConfigSpec
			$spec.host[0].operation  = "add"
			$spec.host[0].host       = New-Object VMware.Vim.ManagedObjectReference
			$spec.host[0].host.type  = "HostSystem"
			$spec.host[0].host.Value = (get-vmhost $vmhost | get-view).MoRef.value
			
			echo "getting vmhostadapter... $vmhost, $vds"
			$arrCount   = [int]$pnicArray.Count
			$pnicDevice = New-Object System.String[] ($arrCount)
			
			Write-Host "Adding ${vmhost} to VDSwitch ${vds}..." -ForegroundColor "${actionClr}"
			Write-Host " "
			Write-Host "Collecting spec data..." -ForegroundColor "blue"
			
			$_this = Get-View -Id "VmwareDistributedVirtualSwitch-${dvsVal}"
			$_this.ReconfigureDvs_Task($spec)
			
			if ($? -eq "True")
			{
				Write-Host "Adding of ${vmhost} to VDSwitch ${vds} was successful!" -ForegroundColor "${successClr}"
				Write-Host " "
				
				disconnectVI
				
				Write-Host "Connecting to host, ${vmhost}..." -ForegroundColor "${actionClr}"
				Connect-VIServer -Server "${vmhost}" -User $hostUser -Password $hostPass -ea SilentlyContinue | Out-Null
				Write-Host "Adding uplinks..." -ForegroundColor "${actionClr}"
				
				$i = 0
				foreach ($pnicVal in $pnicArray)
				{
					Write-Host "Adding ${pnicVal} to uplink port..." -ForegroundColor "${actionClr}"
					Write-Host " "
					$pnicDevice[$i] = $pnicVal
					
					#Remove host physical adapter from VDSwitch "IF" already exists
					get-vmhost -name "${vmhost}" | get-vmhostnetworkadapter -physical -name $pnicVal | remove-vdswitchphysicalnetworkadapter -confirm:$false -ea silentlycontinue
					
					$vmhostNetworkAdapter = Get-VMHost "${vmhost}" | Get-VMHostNetworkAdapter -Physical -Name $pnicVal
					
					#Add host physical adapter back to uplink port
					Get-VDSwitch -Name "${vds}" | Add-VDSwitchPhysicalNetworkAdapter -VMHostNetworkAdapter $vmhostNetworkAdapter -Confirm:$false
					
					if ($? -eq "True")
					{
						Write-Host "Adding of ${pnicVal} to uplink port was successful!" -ForegroundColor "${successClr}"
						Write-Host " "
					} else {
						Write-Host "Adding of ${pnicVal} to uplink port failed for some reason!  Exiting." -ForegroundColor "${errClr}"
						Write-Host " "
						exit
					}
					
					$i++
				}
				
				disconnectVI
			}
			else
			{
				Write-Host "Adding of ${vmhost} to VDSwitch ${vds} failed for some reason!  Exiting." -ForegroundColor "${errClr}"
				Write-Host " "
				exit
			}
		}
		
		function add_VM_NICs()
		{
			#Change to multi-mode vcenter management
			Set-PowerCLIConfiguration -DefaultVIServerMode Multiple -Confirm:$false | Out-Null
			
			connectVI $sourceVI
			connectVI $destVI
			
			Write-Host "Reconfiguring VMs..." -ForegroundColor "${actionClr}"
			Write-Host " "
			get-content $VLANinfo | foreach {
				$VM = $_.ToString().split("`t")[1].trim()
				
				$powerState = (Get-VM -Name "${VM}" -Server $sourceVI).powerstate
				
				Write-Host "Adding ${VM} and associated NIC(s)..." -ForegroundColor "${actionClr}"
				
				$adapterCount = (get-vm "${VM}" -Server $sourceVI | get-networkadapter).NetworkName.Count
				$netName      = @()
				$netAdapt     = @()
				
				echo "TOTAL Adapter Count: ${adapterCount}"
				
				for ($n = 0; $n -lt $adapterCount; $n++)
				{
					$IPv4 = ""
					
					echo "Configuring Adapter # $($n + 1)..."
					
					if ($adapterCount -gt 1)
					{
						$netName += [uri]::EscapeDataString((get-vm "${VM}" -Server $sourceVI | get-networkadapter).NetworkName[$n])
					}
					else
					{
						$netName += [uri]::EscapeDataString((get-vm "${VM}" -Server $sourceVI | get-networkadapter).NetworkName)
					}
					
					$netAdapt += (get-vm "${VM}" -Server $sourceVI | get-networkadapter).Name
					
					$virtualportgroup = Get-VM "${VM}" -Server $sourceVI | get-vdswitch -Name "${vds}" | get-vdportgroup -name "$($netName[$n])"
					$virtualportgroup = [uri]::UnEscapeDataString("${virtualportgroup}")
					
					
					if ($powerState -eq "PoweredOff")
					{
						Get-VM "${VM}" -Server $destVI | get-networkadapter -Name "$($netAdapt[$n])" | Set-NetworkAdapter -networkname $virtualportgroup -Confirm:$false | Set-NetworkAdapter -WakeOnLan:$true -StartConnected:$true -Confirm:$false
					}
					else
					{
						Get-VM "${VM}" -Server $destVI | get-networkadapter -Name "$($netAdapt[$n])" | Set-NetworkAdapter -networkname $virtualportgroup -Confirm:$false | Set-NetworkAdapter -WakeOnLan:$true -StartConnected:$true -Connected:$true -Confirm:$false
					}
					
					if ($? -eq "True")
					{
						Write-Host "Successfully added NIC(s)!" -ForegroundColor "${successClr}"
						Write-Host " "
						
						if ($powerState -ne "PoweredOff")
						{
							(get-vm "${VM}" -Server $destVI) | %{
								$IPv4 = ""
								$vm   = $_
								
								$vm.Guest.Nics | %{
									$vminfo = $_
									$vminfo.Device.Name | ?{ $_ -eq "$($netAdapt[$n])" } | %{ $IPv4 = $vminfo.IPAddress | where { ([IPAddress]$_).AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork } }	
								}
							}
							
							echo "${VM}: Successfully added NIC -- $($netAdapt[$n]) associated with portgroup, ${virtualportgroup}, with IP of ${IPv4}!" | Add-Content $completeVMList
							
							if (test-connection $IPv4 -Quiet -Count 1)
							{
								Write-Host "VM Power State: " -f "${basicClr}" -NoNewline; Write-Host "${powerState} " -f "${menuItemClr}" -NoNewline; Write-Host "--" -f "${basicClr}" -NoNewline; Write-Host " Successfully able to PING ${IPv4}!" -ForegroundColor "${successClr}"
								Write-Host " "
								
								echo "${VM}: Successfully able to PING ${IPv4}!" | Add-Content $completeVMList
							}
							else
							{
								Write-Host "Failed to PING ${IPv4}!" -ForegroundColor "${errClr}"
								Write-Host " "
								
								echo "${VM}: Failed to PING ${IPv4}!" | Add-Content $completeVMList
								echo "${VM}: Failed to PING ${IPv4}!" | Add-Content $failedVMList
							}
						}
						else
						{
							Write-Host "VM Power State: " -f "${basicClr}" -NoNewline; Write-Host "${powerState} " -f "${menuItemClr}" -NoNewline; Write-Host "--" -f "${basicClr}" -NoNewline; Write-Host " NO NEED TO PING!" -ForegroundColor "${powerStateClr}"
							Write-Host " "
							echo "${VM}: is powered off!  There is no IP to PING!" | Add-Content $completeVMList
						}
					}
					else
					{
						Write-Host "Failed to add NIC(s)!" -ForegroundColor "${errClr}"
						Write-Host " "
						
						echo "${VM}: Failed to add NIC -- $($netAdapt[$n]) associated with portgroup, ${virtualportgroup}!" | Add-Content $completeVMList
						echo "${VM}: Failed to add NIC(s)!" | Add-Content $failedVMList
					}
				}
				
				echo " " | Add-Content $completeVMList
			}
			
			if ($global:DefaultVIServers)
			{
				Disconnect-VIServer -Server $global:DefaultVIServers -Force -Confirm:$false
			}
		}
		
		function exportVMNotes()
		{
			$script:vmNoteList = "c:\vm_note_list_${vmhost}_${dateFormat}.csv"
			
			if (test-path -Path "${vmNoteList}")
			{
				Remove-Item "${vmNoteList}"
			}
			
			connectVI $sourceVI
			
			Write-Host "Exporting VM note data..." -ForegroundColor "${actionClr}"
			Write-Host " "
			
			$vmList = Get-VMHost "${sourceHost}" | Get-VM
			$noteList = @()
			foreach ($vm in $vmList)
			{
				$row       = "" | Select Name, Notes
				$row.name  = $vm.Name
				$row.Notes = $vm | select -expandproperty Notes
				$notelist += $row
			}
			$noteList | Export-Csv "${vmNoteList}" –NoTypeInformation
			
			exportVMAnnotations
			exportHOSTAnnotations
			
			disconnectVI
		}
		
		function exportVMAnnotations()
		{
			$script:vmAnnotationsList = "c:\vm_annotations_list_${vmhost}_${dateFormat}.csv"
			
			if (test-path -Path "${vmAnnotationsList}")
			{
				Remove-Item "${vmAnnotationsList}"
			}
			
			Write-Host "Exporting VM annotation data..." -ForegroundColor "${actionClr}"
			Write-Host " "
			
			Get-VMHost "${sourceHost}" | Get-VM | ForEach-Object {
				$VM = $_
				$VM | Get-Annotation |
				ForEach-Object {
					$Report       = "" | Select-Object VM, Name, Value
					$Report.VM    = $VM.Name
					$Report.Name  = $_.Name
					$Report.Value = $_.Value
					$Report
				}
			} | Export-Csv -Path "${vmAnnotationsList}" -NoTypeInformation
		}
		
		function exportHOSTAnnotations()
		{
			$script:hostAnnotationsList = "c:\host_annotations_list_${vmhost}_${dateFormat}.csv"
			
			if (test-path -Path "${hostAnnotationsList}")
			{
				Remove-Item "${hostAnnotationsList}"
			}
			
			Write-Host "Exporting HOST annotation data..." -ForegroundColor "${actionClr}"
			Write-Host " "
			
			Get-VMHost "${sourceHost}" | ForEach-Object {
				$vmHOST = $_
				$vmHOST | Get-Annotation |
				ForEach-Object {
					$Report                 = "" | Select-Object AnnotatedEntity, Name, Value
					$Report.AnnotatedEntity = $vmHOST.AnnotatedEntity
					$Report.Name            = $_.Name
					$Report.Value           = $_.Value
					$Report
				}
			} | Export-Csv -Path "${hostAnnotationsList}" -NoTypeInformation
		}
		
		function importVMNotes()
		{
			if (test-path -Path "${vmNoteList}")
			{
				connectVI $destVI
				Write-Host "Importing VM note data..." -ForegroundColor "${actionClr}"
				Write-Host " "
				
				Get-ChildItem "${vmNoteList}" | ForEach {
					
					$check = Import-Csv $_
					
					if ($check)
					{
						$noteList = Import-Csv "${vmNoteList}"
						foreach ($nLine in $noteList)
						{
							$pattern = "\b$($nLine.Name)\b"
							
							if (!(get-content "${templatesList}" | Select-String -pattern ($pattern) -AllMatches)) 
							{
								if ($nLine.Notes -ne "")
								{
									Set-VM -VM $nLine.Name -Notes $nLine.Notes -Confirm:$false
								}
							}
						}
					}
				}
				
				importVMAnnotations
				importHOSTAnnotations
				
				disconnectVI
			} else {
				Write-Host "No NOTES file!" -ForegroundColor "${errClr}"
				Write-Host " "
			}
		}
		
		function importVMAnnotations()
		{
			if (test-path -Path "${vmAnnotationsList}")
			{
				Write-Host "Importing VM annotation data..." -ForegroundColor "${actionClr}"
				Write-Host " "
				
				Get-ChildItem "${vmAnnotationsList}" | ForEach {
					
					$check = Import-Csv $_
					
					if ($check)
					{
						Import-Csv -Path "${vmAnnotationsList}" | ForEach-Object {
							$pattern = "\b$($_.VM)\b"
							
							if (!(get-content "${templatesList}" | Select-String -pattern ($pattern) -AllMatches))
							{
								New-CustomAttribute -Name $_.Name -TargetType VirtualMachine -ea SilentlyContinue
								Get-VM $_.VM -Server $destVI | Set-Annotation -CustomAttribute $_.Name -Value $_.Value -ea SilentlyContinue
							}
						}
					}
				}
			}
			else
			{
				Write-Host "No VM ANNOTATIONS file!" -ForegroundColor "${errClr}"
				Write-Host " "
			}
		}
		
		function importHOSTAnnotations()
		{
			if (test-path -Path "${hostAnnotationsList}")
			{
				Write-Host "Importing HOST annotation data..." -ForegroundColor "${actionClr}"
				Write-Host " "
				
				Get-ChildItem "${hostAnnotationsList}" | ForEach {
					
					$check = Import-Csv $_
					
					if ($check)
					{
						Import-Csv -Path "${hostAnnotationsList}" | ForEach-Object {
							New-CustomAttribute -Name $_.Name -TargetType VMHost -ea SilentlyContinue
							Get-VMHost "${sourceHost}" | Set-Annotation -CustomAttribute $_.Name -Value $_.Value -ea SilentlyContinue
						}
					}
				}
			}
			else
			{
				Write-Host "No HOST ANNOTATIONS file!" -ForegroundColor "${errClr}"
				Write-Host " "
			}
		}
		
		function convertTemplates_to_VMs()
		{
			Write-Host "Converting templates to VMs..." -ForegroundColor "${actionClr}"
			Write-Host " "
			
			if (Test-Path -Path "${templatesList}")
			{
				Remove-Item -Path "${templatesList}"
			}
			
			$templates = Get-Template -Server $sourceVI -Location "${sourceHost}"
			
			if ($templates.count -gt 0)
			{
				foreach ($template in $templates)
				{
					echo "${template}" | Add-Content "${templatesList}"
					
					Set-Template "${template}" -ToVM -Confirm:$false | Out-Null
					
					if ($? -eq "True")
					{
						Get-VM "${template}" | Move-VM -Destination "${sourcehost}"
						
						if ($? -eq "True")
						{
							Write-Host "Successfully converted ${template} to VM and made sure it is still with same host!" -ForegroundColor "${successClr}"
							Write-Host " "
						}
						else
						{
							Write-Host "Something went wrong with the moving of VM after conversion from template!" -ForegroundColor "${errClr}"
							Write-Host " "
						}
					}
					else
					{
						Write-Host "Something went wrong with the conversion from template to VM!" -ForegroundColor "${errClr}"
						Write-Host " "
					}
				}
			}
		}
		
		function convertVMs_to_Templates()
		{
			Write-Host "Converting VMs back to templates..." -ForegroundColor "${actionClr}"
			Write-Host " "
			
			if (Test-Path -Path "${templatesList}")
			{
				get-content ${templatesList} | foreach {
					$template = $_.ToString().split("`t")[0]
					
					Get-VM "${template}" | Set-VM -ToTemplate -Confirm:$false
					
					if ($? -eq "True")
					{
						Write-Host "Successfully converted ${template} from VM back to template!" -ForegroundColor "${successClr}"
						Write-Host " "
					}
					else
					{
						Write-Host "Something went wrong with the conversion from VM back to template!" -ForegroundColor "${errClr}"
						Write-Host " "
					}
				}
			}
		}
		
		
		# ------- Connect to Source VI -------
		connectVI $sourceVI
		
		# ------- Collect VLAN Information for VMs -------
		getVLANInfo $sourceHost
		
		# ------- Add host to Destination VI -------
		AddHost $sourceDC $sourceHost $VDS
		
		# ------- Export VM note data -------
		exportVMNotes
		
		# ------- Import VM note data -------
		importVMNotes
		
	} else {
		Write-Host "You either chose the wrong file, hit `"Cancel`", or something is wrong with the file!" -ForegroundColor "${errClr}"
		Write-Host " "
	}
	
	if (Test-Path -Path "${failedVMList}")
	{
		Write-Host "There were VMs that failed to migrate from a networking perspective." -ForegroundColor "${actionClr}"
		Write-Host "Opening list in default text editor now..."
		Write-Host " "
		Invoke-Item "${failedVMList}"
	}
	
	if (Test-Path -Path "${completeVMList}")
	{
		Write-Host "Retrieving complete VM status list..." -ForegroundColor "${actionClr}"
		Write-Host "Opening list in default text editor now..."
		Write-Host " "
		Invoke-Item "${completeVMList}"
	}
}

function chkVMAdapters()
{
	if ($global:DefaultVIServers)
	{
		Disconnect-VIServer -Server $global:DefaultVIServers -Force
	}
	
	$script:VMNOPrtGrp = "c:\VM_with_NO_Assigned_PortGroup_${SourceHost}_${dateFormat}.txt"
	
	if (Test-Path -Path "${VMNOPrtGrp}")
	{
		Remove-Item "${VMNOPrtGrp}"
	}
	
	if (promptForFile)
	{
		#Change to multi-mode vcenter management
		Set-PowerCLIConfiguration -DefaultVIServerMode Multiple -Confirm:$false | Out-Null
		
		connectVI $sourceVI
		connectVI $destVI
		
		Get-VMHost "${sourceHost}" -Server "${sourceVI}" | Get-VM -Server "${sourceVI}" | foreach {
			
			$pattern = "\b$($_.Name)\b"
			
			if (!(get-content "${templatesList}" | Select-String -pattern ($pattern) -AllMatches))
			{
				$VM = $_.Name
				
				$getVMNetwork = (get-vm "${VM}" -Server "${destVI}" | get-view).Network
				
				if ($getVMNetwork)
				{
					Write-Host "VM, ${VM}, looks good from a portgroup perspective!" -ForegroundColor "${successClr}"
					
					if ((get-vm "${VM}" -Server "${sourceVI}" | get-view).Network.count -eq (get-networkadapter "${VM}" -Server "${destVI}").Name.count)
					{
						Write-Host "VM, ${VM}, has $((get-networkadapter "${VM}" -Server "${destVI}").Name.count) NIC(s), which should be the same as the source!" -ForegroundColor "${successClr}"
					}
					else
					{
						Write-Host "VM, ${VM}, has $((get-networkadapter "${VM}" -Server "${destVI}").Name.count) NIC(s), which is NOT the same as the source!  Fix immediately!" -ForegroundColor "${errClr}"
						
						for ($p = 0; $p -lt $((get-networkadapter "${VM}" -Server "${sourceVI}").Name.count); $p++)
						{
							echo "${VM}	Should be assigned to portgroup: $((Get-NetworkAdapter "${VM}" -Server "${sourceVI}").NetworkName[$p])" | Add-Content $VMNOPrtGrp
						}
					}
				}
				else
				{
					Write-Host "VM, ${VM}, is NOT assigned to any port group!  Fix immediately!" -ForegroundColor "${errClr}"
					for ($p = 0; $p -lt $((get-networkadapter "${VM}" -Server "${sourceVI}").Name.count); $p++)
					{
						echo "${VM}	Should be assigned to portgroup: $((Get-NetworkAdapter "${VM}" -Server "${sourceVI}").NetworkName[$p])" | Add-Content $VMNOPrtGrp
					}
				}
				
				Write-Host " "
			}
		}
		
		disconnectVI
		
		if (Test-Path -Path "${VMNOPrtGrp}")
		{
			Write-Host "There are some VMs that do not have an associated network adapter." -ForegroundColor "${actionClr}"
			Write-Host "Opening list in default text editor now..."
			Write-Host " "
			Invoke-Item "${VMNOPrtGrp}"
		}
		else
		{
			Write-Host "No problems!  All VMs are assigned to a port group!" -ForegroundColor "${successClr}"
			Write-Host " "
		}
	}
	else
	{
		Write-Host "You either chose the wrong file, hit `"Cancel`", or something is wrong with the file!" -ForegroundColor "${errClr}"
		Write-Host " "
	}
}

function migrateFolders()
{
	$ErrorActionPreference = "silentlycontinue"
	
	Add-PSSnapin VMware.VimAutomation.Core
	if ($global:DefaultVIServers)
	{
		Disconnect-VIServer -Server $global:DefaultVIServers -Force
	}
	
	if (promptForFile)
	{
		connectVI $sourceVI
		
		$roleFile  = "c:\roles.xml"
		$permsFile = "c:\perms.xml"
		$vcFolders = "c:\folders.txt"
		
		if (Test-Path "${roleFile}")
		{
			Remove-Item "${roleFile}"
		}
		
		if (Test-Path "${permsFile}")
		{
			Remove-Item "${permsFile}"
		}
		
		if (Test-Path "${vcFolders}")
		{
			Remove-Item "${vcFolders}"
		}
		
		##EXPORT FOLDERS & PERMISSIONS
		Write-Host "Exporting folder structure and permissions from source, ${sourceVI}..." -ForegroundColor "${actionClr}"
		Write-Host "This is going to take some time, please be patient!" -Foreground "${waitMsgClr}"
		Write-Host " "
		
		get-virole | where { $_.issystem -eq $false } | export-clixml "${roleFile}"
		Get-VIPermission | export-clixml "${permsFile}"
		
		Filter Get-FolderPath {
			$_ | Get-View | % {
				$row      = "" | select Name, Path
				$row.Name = $_.Name
				
				$current  = Get-View $_.Parent
				$path     = $_.Name
				do
				{
					$parent = $current
					if ($parent.Name -ne "vm") { $path = $parent.Name + "\" + $path }
					$current = Get-View $current.Parent
				}
				while ($current.Parent -ne $null)
				$row.Path = $path
				$row
			}
		}
		
		# Export all folders
		$report = @()
		$report = get-datacenter $sourceDC -Server $sourceVI | Get-folder vm | get-folder | Get-Folderpath
		#Replace the top level with vm
		foreach ($line in $report)
		{
			$line.Path = ($line.Path).Replace($sourceDC + "\", "vm\")
			$line.Path = ($line.Path).Replace("\", "/")
			$lineName  = $line.Name
			$linePath  = $line.Path
			echo "${lineName}	${linePath}" | Add-Content "${vcFolders}"
		}
		
		disconnectVI
		
		if (test-path "${vcFolders}")
		{
			Add-PSSnapin VMware.VimAutomation.Core
			if ($global:DefaultVIServers)
			{
				Disconnect-VIServer -Server $global:DefaultVIServers -Force
			}
			
			connectVI $destVI
			
			##IMPORT FOLDERS & PERMISSIONS
			Write-Host "Importing folder structure and permissions to destination, ${destVI}..." -ForegroundColor "${actionClr}"
			Write-Host "This is going to take some time, please be patient!" -Foreground "${waitMsgClr}"
			Write-Host " "
			
			$startFolder   = 'vm'
			$startLocation = Get-Folder -Name $startFolder
			
			foreach ($line in Get-Content "${vcFolders}")
			{
				$folderPath = $line.split("`t")[1]
				$location   = $startLocation
				
				$folderPath.TrimStart('/').Split('/') | where{ $_ -ne 'vm' } | %{
					$tgtFolder = Get-Folder -Name $_ -Location $location[0] -ErrorAction SilentlyContinue
					if (!$tgtFolder)
					{
						$location = New-Folder -Name $_ -Location $location[0]
					}
					else
					{
						$location = $tgtFolder
					}
				}
			}
			
			foreach ($thisRole in (import-clixml "${roleFile}")) { if (!(get-virole $thisRole.name -erroraction silentlycontinue)) { new-virole -name $thisRole.name -Privilege (get-viprivilege -id $thisRole.PrivilegeList -erroraction silentlycontinue) | Out-Null } }
			
			foreach ($thisPerm in (import-clixml "${permsFile}"))
			{
				$permPrincipal = $thisPerm.principal
				get-folder $thisPerm.entity.name -erroraction silentlycontinue | new-vipermission -role $thisPerm.role -Principal "${permPrincipal}" -propagate $thisPerm.Propagate | Out-Null
			}
			
			disconnectVI
		}
		else
		{
			Write-Host "One of the data export files is missing!" -ForegroundColor "${errClr}"
			Write-Host " "
			exit
		}
	}
	else
	{
		Write-Host "You either chose the wrong file, hit `"Cancel`", or something is wrong with the file!" -ForegroundColor "${errClr}"
		Write-Host " "
	}
}

function migrateVMs()
{
	Add-PSSnapin VMware.VimAutomation.Core
	if ($global:DefaultVIServers)
	{
		Disconnect-VIServer -Server $global:DefaultVIServers -Force
	}
	
	if (promptForFile)
	{
		connectVI $sourceVI
		
		$vmLocFile = "c:\vm-locations_${sourceHost}_${dateFormat}.txt"
		
		if (Test-Path "${vmLocFile}")
		{
			Remove-Item "${vmLocFile}"
		}
		
		Write-Host "Exporting VM folder placement from source, ${sourceVI}..." -ForegroundColor "${actionClr}"
		Write-Host " "
		
		filter Get-FolderPath {
			$_ | Get-View | % {
				$row      = "" | select Name, Path
				$row.Name = $_.Name
				
				$current  = Get-View $_.Parent
				$path     = $_.Name
				do
				{
					$parent = $current
					if ($parent.Name -ne "vm") { $path = $parent.Name + "\" + $path }
					$current = Get-View $current.Parent
				}
				while ($current.Parent -ne $null)
				$row.Path = $path
				$row
			}
		}
		
		
		##Export all VM locations
		$report = @()
		
		if (Test-Path -Path "${VLANinfo}")
		{
			get-content ${VLANinfo} | foreach {
				$VM             = $_.ToString().split("`t")[1]
				$NetAdapterName = $_.ToString().split("`t")[2]
				$NetworkName    = $_.ToString().split("`t")[3]
				$MAC            = $_.ToString().split("`t")[5]
				
				$report         = get-datacenter $sourceDC | get-vmhost $sourceHost | get-vm -Name "${VM}" | Get-Folderpath
				
				$vmPattern      = Split-Path $report.path -Leaf
				$vmPattern      = [regex]::escape($vmPattern)
				$fullPath       = $report.path -ireplace "\\${vmPattern}", ""
				$fullPath       = $fullPath -ireplace "\\", "/"
				echo "${VM}	${fullPath}" | Add-Content "${vmLocFile}"
			}
		}
		
		if (Test-Path -Path "${templatesList}")
		{
			get-content ${templatesList} | foreach {
				$template        = $_.ToString().split("`t")[0]
				
				$report          = get-datacenter $sourceDC | get-vmhost $sourceHost | get-vm -Name "${template}" | Get-Folderpath
				
				$templatePattern = Split-Path $report.path -Leaf
				$templatePattern = [regex]::escape($templatePattern)
				$fullPath        = $report.path -ireplace "\\${templatePattern}", ""
				$fullPath        = $fullPath -ireplace "\\", "/"
				echo "${template}	${fullPath}" | Add-Content "${vmLocFile}"
			}
		}
		
		disconnectVI
		
		Add-PSSnapin VMware.VimAutomation.Core
		if ($global:DefaultVIServers)
		{
			Disconnect-VIServer -Server $global:DefaultVIServers -Force
		}
		
		connectVI $destVI
		
		Write-Host "Importing VMs to proper folder structure to destination, ${destVI}..." -ForegroundColor "${actionClr}"
		Write-Host " "
		
		if (Test-Path "${vmLocFile}")
		{
			get-content ${vmLocFile} | foreach {
				$VM         = $_.ToString().split("`t")[0]
				$folderPath = $_.ToString().split("`t")[1]
				
				$folder     = Get-FolderByPath -Path "${folderPath}"
				$pattern    = "\b${VM}\b"
				
				if (!(get-content "${templatesList}" | Select-String -pattern ($pattern) -AllMatches))
				{
					Get-VM -Name "${VM}" | Move-VM -Destination $folder
				}
				else
				{
					Get-Template -Name "${VM}" | Move-Template -Destination $folder
				}
				
				if ($? -eq "True")
				{
					Write-Host "Successfully moved ${VM} to proper folder!" -ForegroundColor "${successClr}"
					Write-Host " "
				}
				else
				{
					Write-Host "Failed to move ${VM} to proper folder!" -ForegroundColor "${errClr}"
					Write-Host " "
				}
			}
			
			disconnectVI
			
		}
		else
		{
			Write-Host "VM location data file doesn't exist!" -ForegroundColor "${errClr}"
			Write-Host " "
			exit
		}
	}
}

[BOOLEAN]$global:xExitSession = $false
function LoadMenuSystem()
{
	[INT]$xMenu1 = 0
	[INT]$xMenu2 = 0
	[BOOLEAN]$xValidSelection = $false
	while ($xMenu1 -lt 1 -or $xMenu1 -gt 5)
	{
		[System.Console]::Clear()
		#Present the Menu Options
		Write-Host "`n`tWeltonWare VMWare Host Migration`n" -BackgroundColor "${WWLabelBackClr}" -ForegroundColor "${WWLabelForeClr}"
		Write-Host "`t`tPlease select which process you would like to run`n" -ForegroundColor "${menuItemClr}"
		Write-Host "`t`t`t1. Folder Migration (w/ Permissions)" -ForegroundColor "${menuItemClr}"
		Write-Host "`t`t`t2. Host Migration" -ForegroundColor "${menuItemClr}"
		Write-Host "`t`t`t3. VM Migration to Folders" -ForegroundColor "${menuItemClr}"
		Write-Host "`t`t`t4. " -ForegroundColor "${menuItemclr}" -NoNewline; Write-Host "POST MIGRATION: " -ForegroundColor "${powerStateClr}" -NoNewline; Write-Host "Check Migrated VMs for Assigned Port Group" -ForegroundColor "${menuItemClr}"
		Write-Host "`t`t`t5. Quit and exit`n" -ForegroundColor "${menuItemClr}"
		#Retrieve the response from the user
		[int]$xMenu1 = Read-Host "`t`tEnter Menu Option Number"
		if ($xMenu1 -lt 1 -or $xMenu1 -gt 5)
		{
			Write-Host "`tPlease select one of the options available.`n" -ForegroundColor "${errClr}"; start-Sleep -Seconds 1
		}
	}
	Switch ($xMenu1)
	{
		#User has selected a valid entry.. load next menu
		1 {
			migrateFolders
		}
		2 {
			HostMigration
		}
		3 {
			migrateVMs
		}
		4 {
			chkVMAdapters
		}
		default { $global:xExitSession = $true; break }
	}
}
LoadMenuSystem
if ($xExitSession)
{
	#exit-pssession -ea SilentlyContinue | Out-Null
	exit
}
else
{
	#.\hostMigration_test.ps1 #Loop the function
}