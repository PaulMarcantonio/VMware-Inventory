<h2>VM Inventory Report</h2>
<strong><u>Description</u>:</strong> 
  <br/>&nbsp;&nbsp;&nbsp;&nbsp;Pulls all VM within multiple VCenters Provided  
  
<strong><u>Useage</u>:</strong> 
  <br/>&nbsp;&nbsp;&nbsp;&nbsp;Can be launched in a Powershell session on a system that has VMware PowerCLI shell also installed on the System. Also great as a task manager task to run on schedule. This provides an Excel file with completed inventory along with statistics.

<strong><u>Programmer</u>:</strong>
     <br/>&nbsp;&nbsp;&nbsp;&nbsp;Paul Marcantonio
     
<strong><u>Date</u>:</strong>
     <br/>&nbsp;&nbsp;&nbsp;&nbsp;June 2 2018
     
<strong><u>Version</u>:</strong>
     <br/>&nbsp;&nbsp;&nbsp;&nbsp;2.2

<strong><u>Pre-Condition</u>:</strong>
  <br/>&nbsp;&nbsp;&nbsp;&nbsp;Powershell installed and with execution policy set to remote signed
  Exchange Powershell Shell (version to match Exchange Server version) is installed on this system as well
  VMware PowerCLI (compatible with VCenter version trying to connect to) must be installed on the server that this script if running on.

<strong><u>Project Title</u>:</strong>
<br/>&nbsp;&nbsp;&nbsp;&nbsp;VCenter VM Inventory report from each VCenter listed

<strong><u>Objective</u>:</strong>
    <br/>&nbsp;&nbsp;&nbsp;&nbsp;Email an Excel file of all VMs (with details) within each VCenter provided
	<ul>
		<li>Obtain encrypted user name and password from text files used for VCenter authentication</li> 
		<li>For each VCenter listed obtain details on each VM and add to array for processing</li>
		<li>Authenticate to VCenter</li>
		<li>Capture all VMs and obtain the following for each:
		<ul>
		<li>General Detail
		<ul>
			<li>VCenter Name</li>
			<li>Power State</li>
			<li>Up Time (Days)</li> 
			<li>Up Time (Hours)</li> 
			<li>Status (Overall VMware Status)</li>
			<li>Number of CPUs</li>
			<li>Number of Cores</li> 
			<li>RAM (GB)</li>
			<li>Total Storage(GB)</li>
			<li>Used Storage(GB)</li> 
			<li>Free Storage(GB)</li> 
			<li>VM FQDN</li>
			<li>IP Address</li> 
			<li>Guest Full Name (VM)</li>
			<li>VMware Tools Health</li> 
			<li>VMware Tools Status</li> 
			<li>VMware Tools Version</li> 
			<li>VM Create Date</li>
			<li>Notes (Admin entered)</li>
			</ul></li>
			<li><strong>Utilization</strong>
			<ul>
				<li>Overall Cpu Demand</li> 
				<li>Overall Cpu Readiness</li>
				<li>Guest Memory Usage </li>
				<li>Host Memory Usage </li>
				<li>Guest Heartbeat Status </li>
				<li>Distributed Cpu Entitlement </li>
				<li>Distributed Memory Entitlement </li>
				<li>Static Cpu Entitlement </li>
				<li>Static Memory Entitlement</li>
				<li>Granted Memory </li>
				<li>Private Memory </li>
				<li>Shared Memory </li>
				<li>Swapped Memory </li>
				<li>Ballooned Memory </li>
				<li>Consumed Overhead Memory</li> 
				<li>FtLog Bandwidth</li> 
				<li>Ft Secondary Latency</li> 
				<li>Ft Latency Status</li> 
				<li>Compressed Memory</li> 
				<li>Ssd Swapped Memory</li>
			</ul></li>
		<li>Add above VM details to the VM master array</li>
		<li>Close connect to VCenter</li>
		<li>Build Excel File and add content from VM master array </li>
		<li>Send email a copy of the Excel file to all listed emails receipents</li>
	</ul>
	<br/>&nbsp;&nbsp;&nbsp;&nbsp;Send Email to team with html table results and HTML file attachement

<strong><u>Pre-Condition(s)</u>:</strong>
     <br/>&nbsp;&nbsp;&nbsp;&nbsp;Powershell Execution Policy set to Remove Signed (Get-ExecutionPolicy, Set-ExecutionPolicy RemoteSigned)
     <br/>&nbsp;&nbsp;&nbsp;&nbsp;Exchange Management Shell installed on the scripting server (Using the Exchange ISO and choose management console only)
     <br/>&nbsp;&nbsp;&nbsp;&nbsp;User this script executes under should have Exchange Enterprise rights (https://docs.microsoft.com/en-us/powershell/exchange/find-exchange-cmdlet-permissions?view=exchange-ps)
     <br/>&nbsp;&nbsp;&nbsp;&nbsp;The server this script runs on must be on the "allow relay" on any and all load balancers infront of exchange and within exchange

<strong><u>Post-Condition(s)</u>:</strong>
     <br/>&nbsp;&nbsp;&nbsp;&nbsp;New HTML file with script data included
     <br/>&nbsp;&nbsp;&nbsp;&nbsp;Email sent to the desired email list with:
          <ul>
            <li>HTML backup summary table included in message body</li>
            <li>Copy of the HTML file with all details from script.</li>
          </ul>
 <strong><u>Installation</u>:</strong>
     <br/>&nbsp;&nbsp;&nbsp;&nbsp;Make sure Powershell version 3 or above is installed on the server (Server Role and Features (Windows Powershell)
     <br/>&nbsp;&nbsp;&nbsp;&nbsp;Domain Administrator Powershell is activated on Server running the script by Roles and Features or running "Import-Module ServerManagerAdd-WindowsFeature RSAT-AD-PowerShell"
     
<strong><u>Contributing</u>:</strong>
    <br/>&nbsp;&nbsp;&nbsp;&nbsp;None

<strong><u>Citations</u>:</strong>
     <br/>&nbsp;&nbsp;&nbsp;&nbsp;None
	 
<strong><u>Output Examples</u>:</strong>
    <br/>&nbsp;&nbsp;&nbsp;&nbsp;Healthy results from script run (Exchange_DB_Check2020-12-04_T16_00_09-Healthy.html)
    <br/>&nbsp;&nbsp;&nbsp;&nbsp;Errors found from script run (Exchange_DB_Check2020-11-20_T14_49_54 Errors.html) 	 
	 
