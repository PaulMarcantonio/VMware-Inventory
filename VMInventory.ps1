<#
Guest VM Inventory Report
Description
  Pulls all VM within multiple VCenters Provided  
  
Usage
  Provides an Email with attached Excel file of all VMs with associated General and Utilization details.

Programmer
     Paul Marcantonio
     
Date
     June 2 2018
     
Version
     2.2

Pre-Condition
  Powershell installed with execution policy set to remote signed
  VMware PowerCLI (compatible with VCenter version connecting to) must be installed on the server that this script runs on.
  EPPlus must be installed and script pointed towards the location of EPPlus DLL (https://www.nuget.org/packages/EPPlus/)

Project Title
VCenter VM Inventory report from each VCenter listed

Objective
    Email an Excel file of all VMs (with details) within each VCenter provided
		Obtain encrypted user name and password from text files used for VCenter authentication 
		For each VCenter listed obtain details on each VM and add to array for processing
		Authenticate to VCenter
		Capture all VMs and obtain the following for each:
		
		General Detail
		
			VCenter Name
			Power State
			Up Time (Days) 
			Up Time (Hours) 
			Status (Overall VMware Status)
			Number of CPUs
			Number of Cores 
			RAM (GB)
			Total Storage(GB)
			Used Storage(GB) 
			Free Storage(GB) 
			VM FQDN
			IP Address 
			Guest Full Name (VM)
			VMware Tools Health 
			VMware Tools Status 
			VMware Tools Version 
			VM Create Date
			Notes (Admin entered)
			
			<strong>Utilization</strong>
			
				Overall Cpu Demand 
				Overall Cpu Readiness
				Guest Memory Usage 
				Host Memory Usage 
				Guest Heartbeat Status 
				Distributed Cpu Entitlement 
				Distributed Memory Entitlement 
				Static Cpu Entitlement 
				Static Memory Entitlement
				Granted Memory 
				Private Memory 
				Shared Memory 
				Swapped Memory 
				Ballooned Memory 
				Consumed Overhead Memory 
				FtLog Bandwidth 
				Ft Secondary Latency 
				Ft Latency Status 
				Compressed Memory 
				Ssd Swapped Memory
			
		Add above VM details to the VM master array
		Close connect to VCenter
		Build Excel File and add content from VM master array 
		Send email a copy of the Excel file to all listed emails receipents
	
	Send Email to team with html table results and HTML file attachement

Pre-Condition(s)
     Powershell Execution Policy set to Remove Signed (Get-ExecutionPolicy, Set-ExecutionPolicy RemoteSigned)
     VMware PowerCLI (version compatible with VCenters connecting to) must be installed on the server running this script
     Need to create encrypted user and password files to use for authenticating to the VCenter (ref Generate_Encrypted_Password_File.ps1 in my repository) 
     The server this script runs on must be on the "allow relay" on any and all load balancers infront of exchange and within exchange

Post-Condition(s)
     New Excel file provided with VM inventory data included
     Email sent to the desired email list with excel file attachment
 
 Installation
     Make sure Powershell version 3 or above is installed on the server (Server Role and Features (Windows Powershell)
     VMware PowerCLI (version compatible with VCenters connecting to) must be installed on the server running this script"
	 You have a user with proper credentials to access inventory within the VCenter (test with GUI first)"
	 Download and install EPPlus to create and format the excel file with data (https://www.nuget.org/packages/EPPlus/)."
     
Contributing
    https://www.nuget.org/packages/EPPlus/

Citations
     None
	 
Output Examples
    Excel File from script run (VMware_Inventory-D2020-12-27T07-15-25.xlsx) 

#>

<#
.SYNOPSIS => Initalize global variables, Excel and HTML files, target VCenters, get times stamps, disable VCenter cert checks. 
.PARAMETER
     None
.INPUTS
     Excel file name and location
	 HTML file name and location (currently not used)
.OUTPUTS
    None
#>
FUNCTION INITIALIZE()
{
	#OBTAIN SCRIPT SERVERNAME AND TIME STAMP
		$global:serverName = $env:COMPUTERNAME
		$global:timeStamp = Get-Date -UFormat "D%Y-%m-%dT%H-%M-%S"
	#EXCEL INFORMATION	
		$global:ExcelFile = "\\1168-POWERADMIN\Web Pages\VMware\Inventory\VMware_Inventory-$timeStamp.xlsx"
		$global:ExcelVMs = @()
	#BUILD HTML VARIABLES
		$global:html = "\\1168-POWERADMIN\Web Pages\VMware\Inventory\VMInventoryReport-$timeStamp.html"
		$global:htmlTableBlock = @()
		$global:htmlHeader = @()
		$global:timeStamp = get-date -uformat "%Y-%m-%d at %H:%M:%S"	#Time stamp for use for HTML file naming
	#BUILD VCENTER DETAILS
		$global:COCCAvon2 = "Avl-vcenter-p02.client.cocci.com"
		$global:COCCAvon3 = "Avl-vcenter-p03.client.cocci.com"
		$global:COCCAvon5 = "Avl-vcenter-p05.client.cocci.com"
		$global:COCCValleyForge2 = "vvl-vcenter-p02.client.cocci.com"
		$global:COCCValleyForge3 = "vvl-vcenter-p03.client.cocci.com"
		$global:MetroChelsea = "ChelseaVCenter"
		$global:MetroRentsys = "RentsysVCenter"
	#BUILD VCENTER ARRAYS
		$global:COCCAvonVCenters = @($global:COCCAvon2,$global:COCCAvon3,$global:COCCAvon5)
		$global:COCCValleyForgeVCenters = @($global:COCCValleyForge2,$global:COCCValleyForge3)
		$global:MetroVCenters = @($global:MetroChelsea, $global:MetroRentsys)
		#$global:AllVCenters = @($global:COCCAvonVCenters,$global:COCCValleyForgeVCenters,$global:MetroVCenters)
		$global:AllVCenters = @($global:COCCAvon2, $global:COCCAvon3, $global:COCCAvon5, $global:COCCValleyForge2, $global:COCCValleyForge3, $global:MetroChelsea, $global:MetroRentsys)
		#$global:AllVCenters = @($global:COCCValleyForge)
	#IGNORE CERTIFICATION CHECKS WHEN AUTHENTICATING TO THE VCENTERS
		Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false	#IGNORE INVALID CERTS CHECK
}


<#
.SYNOPSIS => Open a connection to the VCenter (FQDN passed in)
.PARAMETER
    FQDN of the target VCenter
.INPUTS
	Path to each Target VCenter encrypted password file
.OUTPUTS
    Open connections to passed in VCenter
#>
FUNCTION CONNECT_TO_VCENTER ($vc)
{
	#METRO DATA CENTERS READ SECURE PASSWORD FILE AND OPEN CONNECTION TO VCENTER 
		if (($vc -eq $global:MetroChelsea) -or ($vc -eq $global:MetroRentsys))
		{
			$Metrouser = "metrocu\mcupowerscript"
			$Metropassword = Get-Content "e:\MetroITScripts\VMware\Authentication\Metro-password.txt" | ConvertTo-SecureString -Key (Get-Content "e:\MetroITScripts\VMware\Authentication\Metro-aes.key")
			$credentials = New-Object System.Management.Automation.PsCredential($Metrouser, $Metropassword)
			Connect-VIServer -Server $vc -Credential $credentials
		}
	#COCC DATA CENTER READ SECURE PASSWORD FILE AND OPEN CONNECTION TO VCENTER
		else 
		{
			$COCCuser = Get-Content "e:\MetroITScripts\VMware\Authentication\COCCU"
			$COCCpassword = Get-Content "e:\MetroITScripts\VMware\Authentication\COCC-password.txt" | ConvertTo-SecureString -Key (Get-Content "e:\MetroITScripts\VMware\Authentication\COCC-aes.key")
			$credentials = New-Object System.Management.Automation.PsCredential($COCCuser, $COCCpassword)
			Connect-VIServer -Server $vc -Credential $credentials
		}
}

<#
.SYNOPSIS => Force Disconnect/Close from passed in FQDN VCenter without confirmation
.PARAMETER
     FQDN name of the VCenter
.INPUTS
     None
.OUTPUTS
    None
#>
FUNCTION DISCONNECT_FROM_VCENTER ($vc)
{
	Disconnect-VIServer -Server $vc -Force -Confirm:$false
}

<#
.SYNOPSIS => Obtain all VMs from within VCenter FQDN passed in 
.PARAMETER
     FQDN VCenter passed in 
.INPUTS
     None
.OUTPUTS
    Object containing all VMs obtained from the passed in VCenter
#>
FUNCTION GET_VMS ($vcenter)
{
	$Vms = Get-VM -Server $vcenter
	RETURN $Vms
}

<#
.SYNOPSIS => Return an array of the following details associated with each VM with VMs passed and the VCenter
	Associated Information about the VM
	VCenterName (hosting the VM)
	Name (within the VCenter; maybe be the OS Host Notes)
	PowerState
	Up Time (Days)
	Up Time (Hours)
	Health Status
	Number of CPU
	Number of Cores
	RAM MemoryGB)
	Total Storage(GB)
	Used Storage(GB)
	Free(GB)
	guestFQDN (VM FQDN)
	ip
	guest Full Name
	VMware tools Health Status
	VMware tools Status
	VMware tools Version
	create Date (of VM)
	notes (admin Notes)
	Overall CPU Demand
	Overall CPU Readiness
	Guest Memory Usage
	Host Memory Usage (Host VM is hosted on) 
	Guest Heart beat Status 
	Distributed Cpu Entitlement 
	Distributed Memory Entitlement 
	Static CPU Entitlement 
	Static Memory Entitlement 
	Granted Memory 
	Private Memory
	Shared Memory 
	Swapped Memory 
	Ballooned Memory 
	Consumed Overhead Memory
	FtLog Bandwidth 
	FtSecondary Latency 
	FtLatency Status 
	Compressed Memory 
	SsdSwapped Memory

.PARAMETER
    $vms => Array of vms passed in to be processed
	$vcenterName => FQDN of VCENTER associated with VMs passed
.INPUTS
    $global:ExcelVMs => used to add newly created object to
.OUTPUTS
    Update to the $global:ExcelVMs
#>
FUNCTION ADD_TO_EXCEL_ARRAY ($Vms, $vcenterName)
{
	foreach ($vm in $VMs)
	{
		$vmObject = [PSCustomObject]@{
			VCenter = $vcenterName
			Name = $vm.Name
			powerState = $vm.PowerState
			upTimeDays = [timespan]::fromseconds($vm.ExtensionData.Summary.QuickStats.UptimeSeconds).Days 
			upTimeHours = [timespan]::fromseconds($vm.ExtensionData.Summary.QuickStats.UptimeSeconds).Hours
			status = $vm.ExtensionData.OverallStatus
			numCPU = [int]($vm.NumCpu)
			numCores = [int]($vm.CoresPersocket)
			RAM = [int]($vm.MemoryGB)
			'TotalStorage(GB)' = [int]([Math]::Round($vm.ProvisionedSpaceGB,2))
			'UsedStorage(GB)' = [int]([Math]::Round($vm.UsedSpaceGB,2))
			'Free(GB)' = [int]([Math]::Round(($vm.ProvisionedSpaceGB - $vm.UsedSpaceGB), 2))
			guestFQDN = $vm.ExtensionData.Guest.HostName
			ip = $vm.ExtensionData.Guest.IPaddress
			guestFullName = $vm.ExtensionData.Guest.GuestFullName
			toolsHealth = $vm.ExtensionData.Guest.ToolsStatus
			toolsStatus = $vm.ExtensionData.Guest.ToolsRunningStatus
			toolsVersion = $vm.ExtensionData.Guest.ToolsVersion
			createDate = [string]$vm.CreateDate
			notes = $vm.Notes
			OverallCpuDemand = $vm.extensiondata.summary.QuickStats.OverallCpuDemand
			OverallCpuReadiness = $vm.extensiondata.summary.QuickStats.OverallCpuReadiness
			GuestMemoryUsage = $vm.extensiondata.summary.QuickStats.GuestMemoryUsage
			HostMemoryUsage = $vm.extensiondata.summary.QuickStats.HostMemoryUsage
			GuestHeartbeatStatus = $vm.extensiondata.summary.QuickStats.GuestHeartbeatStatus
			DistributedCpuEntitlement = $vm.extensiondata.summary.QuickStats.DistributedCpuEntitlement
			DistributedMemoryEntitlement = $vm.extensiondata.summary.QuickStats.DistributedMemoryEntitlement
			StaticCpuEntitlement = $vm.extensiondata.summary.QuickStats.StaticCpuEntitlement
			StaticMemoryEntitlement = $vm.extensiondata.summary.QuickStats.StaticMemoryEntitlement
			GrantedMemory = $vm.extensiondata.summary.QuickStats.GrantedMemory
			PrivateMemory = $vm.extensiondata.summary.QuickStats.PrivateMemory
			SharedMemory = $vm.extensiondata.summary.QuickStats.SharedMemory
			SwappedMemory = $vm.extensiondata.summary.QuickStatsSwappedMemory
			BalloonedMemory = $vm.extensiondata.summary.QuickStats.BalloonedMemory
			ConsumedOverheadMemory = $vm.extensiondata.summary.QuickStats.ConsumedOverheadMemory
			FtLogBandwidth = $vm.extensiondata.summary.QuickStats.FtLogBandwidth
			FtSecondaryLatency = $vm.extensiondata.summary.QuickStats.FtSecondaryLatency
			FtLatencyStatus = $vm.extensiondata.summary.QuickStats.FtLatencyStatus
			CompressedMemory = $vm.extensiondata.summary.QuickStats.CompressedMemory
			SsdSwappedMemory = $vm.extensiondata.summary.QuickStats.SsdSwappedMemory
		}
		$global:ExcelVMs += $vmObject
	}
	}

<#
.SYNOPSIS => Delete the file associated with the $global:html file
.PARAMETER
    None
.INPUTS
    $global:html => file variable that this will delete
.OUTPUTS
    None
#>
FUNCTION DELETE_HTML_FILE
{
	if (Test-path $global:html)
	{
		Remove-Item -LiteralPath $global:html -Force
	}
}

<#
.SYNOPSIS => Updates the global HTML table block variable with VCenter name and VM details associated with each in VM array. This html block will be used to populate the html file with details. 
.PARAMETER
    $vc = string value of the VCenter Name
	$VMs = array of VM objects that contain the details associated with each VM
.INPUTS
     $global:htmlTableBlock => global variable to hold the html table structure of each vm 
.OUTPUTS
    An updated global:htmlTableBlock
#>
FUNCTION ADD_TO_HTML_BLOCK ($vc, $VMs)
{
	#ADD SNAPS TO HTML BLOCK
		$counter = 0
		foreach ($vm in $VMs)
		{
			$vmName = $vm.Name
			$powerState = $vm.PowerState
			$upDays = [timespan]::fromseconds($vm.ExtensionData.Summary.QuickStats.UptimeSeconds).Days
			$upHours =[timespan]::fromseconds($vm.ExtensionData.Summary.QuickStats.UptimeSeconds).Hours
			$status = $vm.ExtensionData.OverallStatus
			$numCPU = $vm.NumCpu
			$numCores = $vm.CoresPersocket
			$ram = $vm.MemoryGB
			$totalStorage = [Math]::Round($vm.ProvisionedSpaceGB,2)
			$usedStorage = [Math]::Round($vm.UsedSpaceGB,2)
			$free = [Math]::Round(($vm.ProvisionedSpaceGB - $vm.UsedSpaceGB), 2)
			$guestFQDN = $vm.ExtensionData.Guest.HostName
			$ip = $vm.ExtensionData.Guest.IPaddress
			$guestFullName = $vm.ExtensionData.Guest.GuestFullName
			$toolsHealth = $vm.ExtensionData.Guest.ToolsStatus
			$toolsStatus = $vm.ExtensionData.Guest.ToolsRunningStatus
			$toolsVersion = $vm.ExtensionData.Guest.ToolsVersion
			$createDate = [string]$vm.CreateDate
			$notes = $vm.Notes
			#quickstats.balloonedmemory
			
			$counter ++
			$sn = $vm.SnapName
			if ( $VMs -eq $null)
			{
				$global:htmlTableBlock += '<tr><td>'+ $vc +'</td><td colspan="6" align="center">'+ "NO GUEST(VMs)SYSTEMS CAPTURED DURING THIS SCRIPT RUN"+'</td></tr>'
			}
			else
			{
				if ($counter % 2 -eq 0)
				{
						$global:htmlTableBlock += '<button type="button" class="collapsible"><span class="vmname">'+$vmName+'</span><span class="guestFQDN">'+$guestFQDN+'</span><span class="powerState">'+$powerState+'</span><span class="upTime">'+$upDays+' Days '+$upHours+' Hours</span><span class="status">'+$status+'</span><span class="toolsHealth">'+$toolsHealth+'</span><span class="toolsStatus">'+$toolsStatus+'</span><span class="toolsVersion">'+$toolsVersion+'</span></span></a></button>'
						$global:htmlTableBlock += '<div class="content">'
						
						$global:htmlTableBlock += '<ul>'
						$global:htmlTableBlock += '<li><span class="description">IP Address:</span><span class="results">'+$ip+'</span></li>'
						$global:htmlTableBlock += '<li><span class="description">Guest OS Version:</span><span class="results">'+$guestFullName+'</span></li>'
						$global:htmlTableBlock += '<li><span class="description">Number of CPU:</span><span class="results">'+$numCPU+'</span></li>'
						$global:htmlTableBlock += '<li><span class="description">Number of Cores/CPU:</span><span class="results">'+$numCores+'</span></li>'
						$global:htmlTableBlock += '<li><span class="description">RAM (GB)</span><span class="results">'+$ram+'</span></li>'
						$global:htmlTableBlock += '<li><span class="description">Storage Provisioned:</span><span class="results">'+$totalStorage+'</span></li>'
						$global:htmlTableBlock += '<li><span class="description">Storage Used:</span><span class="results">'+$usedStorage+'</span></li>'
						$global:htmlTableBlock += '<li><span class="description">Storage Available:</span><span class="results">'+($totalStorage - $usedStorage)+'</span></li>'+"`n"
						$global:htmlTableBlock += '<li><span class="description">VMware Tools Version:</span><span class="results">'+$toolsVersion+'</span></li>'+"`n"
						$global:htmlTableBlock += '<li><span class="description">VM Create Date:</span><span class="results">'+$createDate+'</span></li>'+"`n"
						$global:htmlTableBlock += '<li><span class="description">VCenter VM Notes:</span><span class="results">'+$notes+'</span></li>'+"`n"
						$global:htmlTableBlock += '</ul>'+"`n"
				        $global:htmlTableBlock += '</div>'+"`n"
				}
				else
				{
					$global:htmlTableBlock += '<button type="button" class="collapsible"><span class="vmname">'+$vmName+'</span><span class="guestFQDN">'+$guestFQDN+'</span><span class="powerState">'+$powerState+'</span><span class="upTime">'+$upDays+' Days '+$upHours+' Hours</span><span class="status">'+$status+'</span><span class="toolsHealth">'+$toolsHealth+'</span><span class="toolsStatus">'+$toolsStatus+'</span><span class="toolsVersion">'+$toolsVersion+'</span></span></a></button>'
						$global:htmlTableBlock += '<div class="content">'
						
						$global:htmlTableBlock += '<ul>'
						$global:htmlTableBlock += '<li><span class="description">IP Address:</span><span class="results">'+$ip+'</span></li>'
						$global:htmlTableBlock += '<li><span class="description">Guest OS Version:</span><span class="results">'+$guestFullName+'</span></li>'
						$global:htmlTableBlock += '<li><span class="description">Number of CPU:</span><span class="results">'+$numCPU+'</span></li>'
						$global:htmlTableBlock += '<li><span class="description">Number of Cores/CPU:</span><span class="results">'+$numCores+'</span></li>'
						$global:htmlTableBlock += '<li><span class="description">RAM (GB)</span><span class="results">'+$ram+'</span></li>'
						$global:htmlTableBlock += '<li><span class="description">Storage Provisioned:</span><span class="results">'+$totalStorage+'</span></li>'
						$global:htmlTableBlock += '<li><span class="description">Storage Used:</span><span class="results">'+$usedStorage+'</span></li>'
						$global:htmlTableBlock += '<li><span class="description">Storage Available:</span><span class="results">'+($totalStorage - $usedStorage)+'</span></li>'+"`n"
						$global:htmlTableBlock += '<li><span class="description">VMware Tools Version:</span><span class="results">'+$toolsVersion+'</span></li>'+"`n"
						$global:htmlTableBlock += '<li><span class="description">VM Create Date:</span><span class="results">'+$createDate+'</span></li>'+"`n"
						$global:htmlTableBlock += '<li><span class="description">VCenter VM Notes:</span><span class="results">'+$notes+'</span></li>'+"`n"
						$global:htmlTableBlock += '</ul>'+"`n"
				        $global:htmlTableBlock += '</div>'+"`n"
				}
			}
		}
}

<#
.SYNOPSIS => Construct the top portion of the HTML file (Header, Title, Scripts, Style) to be later added to file
.PARAMETER
    None Needed
.INPUTS
    $global:htmlHeader => Array that holds the html header details in html form
.OUTPUTS
    Updated global:htmlHeader variable with standard w3.org html file header detail
#>
FUNCTION BUILD_HTML_HEADER ()
{
	$global:htmlHeader += '<!DOCTYPE html>'+"`n"
	$global:htmlHeader += '<html xmlns="http://www.w3.org/1999/xhtml">'+"`n"
	$global:htmlHeader += '<head>'+"`n"
	$global:htmlHeader += '<title>VM Inventory Report</title>'+"`n"
	#$global:htmlHeader += '<script src="../Scripts/sorttable.js"></script>'
	$global:htmlHeader += '<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.5.0/css/bootstrap.min.css">'+"`n"
  	$global:htmlHeader += '<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script>'+"`n"
  	$global:htmlHeader += '<script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.16.0/umd/popper.min.js"></script>'+"`n"
  	$global:htmlHeader += '<script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.5.0/js/bootstrap.min.js"></script>'+"`n"
	$global:htmlHeader += '</head>'+"`n"
			
	$global:htmlHeader += '<style>
		div.container {width: 100%; border: 1px solid gray; background-color:#b5dcb3; margin: auto;}
		div.header {background-color:#b5dcb3; border-bottom: 2px solid black;}
		div.mainSection {background-color:#D4DFDA; height:750px; width:100%;color: #000000;}
		div.footer {background-color:#b5dcb3; height:auto; width:100%;}
			header, footer {padding: 1em; color: #000000; clear: left;text-align: center;}
			caption {font-weight: bold; font-size: 18px; padding-top: 25px; padding-bottom: 25px;}
			th { border-bottom: 2px solid black; padding-left: 20px; padding-right: 20px;}
					td { padding-left: 20px; pading-right: 20px;}
					tr.even {background-color: #f2f2f2;}
					tr.odd {background-color: #c1c1c1;}
					.snap { margin: 0 auto;}
					h3 {text-align: center;}
										
					span.vmname {width: 12%; background-color: red;}
					span.guestFQDN  {width: 12%; background-color: blue;}
					span.powerState  {width: 12%; background-color: red;}
					span.upTime  {width: 12%; background-color: blue;}
					span.status  {width: 12%; background-color: red;}
					span.toolsHealth  {width: 12%; background-color: blue;}
					span.toolsStatus  {width: 12%; background-color: red;}
					span.toolsVersion  {width: 12%; background-color: blue;}
					
					span.description {
						display: inline-block;
						width: 25%;
						padding: 5px;
						border: 1px solid blue;    
						background-color: yellow;
					}
					span.results {
						display: inline-block;
						width: 75%;
						padding: 5px;
						border: 1px solid blue;    
						background-color: red;
					}
					</style>'
			$global:htmlHeader += '<body>'+"`n"
			$global:htmlHeader += '<div class="container">
				<div class="header">
				<header><H1 align="center">Metro & COCC DataCenter VMWARE Daily VM Inventory Report</H1><H2 align="center">This report was last run on '+"$global:timeStamp"+'</H2>'+"`n"
			$global:htmlHeader += '<p align="center">The Powershell <a href="file:///\\'+$global:ServerName+'\Web Pages\VMware\Inventory\VMInventoryReport.html"> Vmware Snap Shot Report</a> runs daily on '+$global:serverName+'.</br>
								  	The script name is located <a href="file:///\\'+$global:ServerName+'\e$\MetroITScripts\VMware\Inventory">here</a> and titled <b>VMwareSnapShotReport</b>. 
								   </p></header>
								   </div>'

}	

<#
.SYNOPSIS => Loads EPPLus.dll (used to build excel file), Clear old excel file, create new excel file,  
.PARAMETER
    $DLLPath => path the EPPlus.dll
.INPUTS
    Path to EPPLUS DLL (I.E "E:\MetroITScripts\Bin\DotNet4\EPPlus.dll")
.OUTPUTS
    None
#>
FUNCTION BUILD_EXCEL_FILE ($Vms)
	{
		#BUILD EXCEL FILE
			# Load EPPlus
				$DLLPath = "E:\MetroITScripts\Bin\DotNet4\EPPlus.dll"
				[Reflection.Assembly]::LoadFile($DLLPath) | Out-Null
			# Clear old File
				#clearCurrentFile $global:ExcelFile
			# Create Excel File
				$ExcelPackage = New-Object OfficeOpenXml.ExcelPackage 
			#Build excel worksheet
				$Worksheet = $ExcelPackage.Workbook.Worksheets.Add("Inventory")
			#Convert data to CSV format 	
				$ProcessesString = $Vms | ConvertTo-Csv -NoTypeInformation | Out-String
				$Format = New-object -TypeName OfficeOpenXml.ExcelTextFormat -Property @{TextQualifier = '"'}
			#Add data to worksheet
				$null=$Worksheet.Cells.LoadFromText($ProcessesString,$Format)
			#Save Excel file after adding data to worksheet	
				$ExcelPackage.SaveAs($global:ExcelFile)
	}

<#
.SYNOPSIS => Setup Email server, Email Receipents, Subject & Body, Send email in HTML format with Excel file to receipents.
.PARAMETER
     None
.INPUTS
    $smtpServer = FQDN to the SMTP Exchange Email DAG
	$emailFrom = SMTP address you want the email to display it is coming from (this can be a fake but must be in proper format) 
	$emailTo = array of valid email addresses you want to send the email to
	$subjectTextSuccess = subject line text showing server name and time stamp 
	$bodyText = html block that will be inserted into email body as HTML
	$global:ExcelFile = attachment that will be included in the email 
.OUTPUTS
    None
#>
FUNCTION SENDEMAIL ()
{
	#Email Server
		$smtpServer = "owa.company.org"													#Metro Email server used to process email
		$smtp = New-Object Net.Mail.SmtpClient($smtpServer)								#

	#Email Addresses
		$emailFrom = "Daily_VM_Inventory_Report_"+$global:ServerName+"@company.org"				#from email address shown in the email
		$emailTo = @("pmarcantonio@company.org","user2@company.org","user3@company.org") #success email address that will recieve logs
		
	#Email Subject & Body Content 
		$subjectTextSuccess = "Daily VMware Inventory Report Script completed on server "+ $global:serverName +" at "+ $global:timeStamp	#subject text for success messages
		$bodyText = $global:htmlHeader + $global:htmlTableBlock

	#Send emails out to the entries in array
		foreach ($rcp in $emailTo)
		{
			Send-MailMessage -from "$emailFrom" -to "$rcp" -subject "$subjectTextSuccess" -body "$bodyText" -BodyAsHtml -Attachments $global:ExcelFile -smtpServer "$smtpServer"
		}
}

<#
.SYNOPSIS => PROGRAM LAUNCH
	Initalize
	Connect to VCenters
	Gather VMs
	Add VM details to array 
	Add array to Excel
	Disconnect from VCenters
	Email Excel to each receipent 
.PARAMETER
     None
.INPUTS
     None
.OUTPUTS
    None
#>
FUNCTION MAIN ()
{
	#INITIALIZE 
		INITIALIZE 
	#OPEN HTML CAPTION AND TABLE HEADERS
		$global:htmlTableBlock += '<div class="mainSection">'
		$global:htmlTableBlock += '<table id="accordion">'
    	$global:htmlTableBlock += '<tr class="card">'

	#CONNECT TO EACH VCENTER AND GET ALL VMs 
		foreach ($vc in $global:AllVCenters)
		{
				Write-Host $vc
			#Connect to VC Center
				CONNECT_TO_VCENTER $vc
			#OBTAIN VMS WITHIN ABOVE CONNECTED VCenter
				$totalVMs = GET_VMs $vc
			#ADD VMs AND VCenter DETAIL TO THE EXCEL ARRAY	
				ADD_TO_EXCEL_ARRAY $totalVMs $vc
			#CLOSE CONNECTION TO TARGET VCenter 
				DISCONNECT_FROM_VCENTER $vc
		}
	#BUILD EXCEL FILE USING THE GLOBAL EXCEL ARRAY CONTAINING ALL VMs CAPTURED FROM ABOVE LOOP
		BUILD_EXCEL_FILE $global:ExcelVMs
	#SEND EMAIL WITH EXCEL RESULTS
		sendEmail
}
#CLEAR SCREEN
	CLS
#RUN PROGRAM
	MAIN