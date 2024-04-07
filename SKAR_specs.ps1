<#
COPYRIGHT SKAR_specs 2024 (C)

Gathers information about computer hardware and software. Outputs to a HTML file.
#>

$PCName = $(hostname)
$FilePath = "$PSScriptRoot\$PCName - $env:UserName - $((Get-Date).tostring("yyyy-MM-dd hh-mm-ss")).tmp"

IF (![System.IO.File]::Exists("$FilePath"))
{
	Add-Content -Path $FilePath -Value '<!DOCTYPE html><html>
	<head>'
	Add-Content -Path $FilePath -Value '<style type="text/css">'
	Add-Content -Path $FilePath -Value '.tg  { border-collapse:collapse; border-color:#9ABAD9; border-spacing:0; margin:0px auto; }'
	Add-Content -Path $FilePath -Value '.tg td{
		background-color:#EBF5FF; border-color:#9ABAD9; border-style:solid; border-width:1px; color:#444;
		font-family:Arial, sans-serif; font-size:14px; overflow:hidden; padding:10px 5px; word-break:normal;
	}'
	Add-Content -Path $FilePath -Value '.tg th{
		background-color:#409cff; border-color:#9ABAD9; border-style:solid; border-width:1px; color:#fff;
		font-family:Arial, sans-serif; font-size:14px; font-weight:normal; overflow:hidden; padding:10px 5px; word-break:normal;
	}'
	Add-Content -Path $FilePath -Value '.tg .tg-0lax{ text-align:left; vertical-align:top }
	@media screen and (max-width: 767px) { .tg { width: auto !important; }.tg col { width: auto !important; }.tg -wrap { overflow-x: auto; - webkit-overflow-scrolling: touch; margin: auto 0px; } }'
	Add-Content -Path $FilePath -Value '</style></head><body>'
}

$BaseBoardInfo = Get-CimInstance -Query "Select Manufacturer, Product from Win32_BaseBoard"
$BIOSInfo = Get-CimInstance Win32_Bios
$CIM_ComputerSystem = Get-CimInstance -ClassName CIM_ComputerSystem
$CIM_BIOSElement = Get-CimInstance -ClassName CIM_BIOSElement
$CIM_OperatingSystem = Get-CimInstance -ClassName CIM_OperatingSystem
$CIM_Processor = Get-CimInstance -ClassName CIM_Processor
$CIM_LogicalDisk = Get-CimInstance -ClassName CIM_LogicalDisk | Where-Object { $_.Name -eq $CIM_OperatingSystem.SystemDrive }


Add-Content -Path $FilePath -Value '<div class="tg-wrap"><table class="tg"><thead>'

#---SECTION START
Add-Content -Path $FilePath -Value '
<tr>
<th class="tg-0lax" colspan="2">'
$PCName | Add-Content -Path $FilePath
Get-Date -Format "dd/MM/yy HH:mm" | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value ' - NAVIGATION</th>
</tr>
</thead>
<tbody>
<tr>'
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax" colspan="2">[<a href="#COMPUTER">COMPUTER</a>] [<a href="#OPERATINGSYSTEM">OPERATING SYSTEM</a>] [<a href="#Motherboard">Motherboard</a>] [<a href="#MEMORY">MEMORY</a>] [<a href="#HARDDRIVES">HARD DRIVES</a>] [<a href="#DISPLAY">DISPLAY</a>] [<a href="#Network">NETWORK</a>] [<a href="#NetworkShares">Network Shares</a>] [<a href="#Networkmappeddrives">Network mapped drives</a>] [<a href="#PRINTERS">PRINTERS</a>] [<a href="#SOFTWARE">SOFTWARE</a>]'
Add-Content -Path $FilePath -Value '</td></tr>'
#row end

#---SECTION START
Add-Content -Path $FilePath -Value '
<tr>
<th class="tg-0lax" colspan="2" id="COMPUTER">COMPUTER</th>
</tr>
</thead>
<tbody>
<tr>'

<#
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax">------------</td>
<td class="tg-0lax">'
COMMAND
Add-Content -Path $FilePath -Value '</td></tr>'
#row end
#>

#row start
Add-Content -Path $FilePath -Value '<td class="tg-0lax">Computer name:</td><td class="tg-0lax">'
"$($env:COMPUTERNAME)" | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</td></tr>'
#row end
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax">Manufacturer:</td>
<td class="tg-0lax">'
"$($CIM_ComputerSystem.Manufacturer)" | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</td></tr>'
#row end
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax">Model:</td>
<td class="tg-0lax">'
"$($CIM_ComputerSystem.Model)" | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</td></tr>'
#row end
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax">Serial Number:</td>
<td class="tg-0lax">'
"$($CIM_BIOSElement.SerialNumber)" | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</td></tr>'
#row end
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax">Processor:</td>
<td class="tg-0lax">'
Get-CimInstance -ClassName Win32_Processor | Select-Object -ExpandProperty Name | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</td></tr>'
#row end
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax">RAM:</td>
<td class="tg-0lax">'
$TotalRAM = (Get-CimInstance -ClassName Win32_ComputerSystem).TotalPhysicalMemory / 1GB; $RAMType = (Get-CimInstance -ClassName Win32_PhysicalMemory | Select-Object -ExpandProperty MemoryType) | ForEach-Object {switch($_) {20 {'DDR'} 21 {'DDR2'} 22 {'DDR3'} 24 {'DDR4'} 26 {'DDR5'} default {'Unknown'}}}
"$TotalRAM GB, RAM Type: $RAMType" | Add-Content -Path $FilePath

#(Get-CimInstance -ClassName Win32_ComputerSystem).TotalPhysicalMemory / 1GB | Add-Content -Path $FilePath
#Get-CimInstance -ClassName Win32_PhysicalMemory | Select-Object -ExpandProperty MemoryType | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</td></tr>'
#row end
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax">Network Devices</td>
<td class="tg-0lax">'
Get-NetAdapter -Physical | Select-Object Name, MacAddress, Status | ConvertTo-Html -Fragment | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</td></tr>'
#row end
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax">Current User:</td>
<td class="tg-0lax">'
"$($env:UserName)" | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</td></tr>'
#row end
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax">Computer Domain</td>
<td class="tg-0lax">'
"$($env:UserDomain)" | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</td></tr>'
#row end
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax">Full Qualified Domain Name</td>
<td class="tg-0lax">'
"$($env:computername+'.'+$($env:userdnsdomain))" | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</td></tr>'
#row end
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax">Domain (True) / Workgroup (false)</td>
<td class="tg-0lax">'
"$((Get-CimInstance -Class Win32_ComputerSystem).PartOfDomain)" | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</td></tr>'
#row end
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax">Workgroup:</td>
<td class="tg-0lax">'
"$((Get-CimInstance -Class Win32_ComputerSystem).Workgroup)"| Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</td></tr>'
#row end

#---SECTION START
Add-Content -Path $FilePath -Value '
	<tr>
	<th class="tg-0lax" colspan="2" id="OPERATINGSYSTEM">OPERATING SYSTEM</th>
	</tr>
	</thead>
	<tbody>
	<tr>'
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax">OS:</td>
<td class="tg-0lax">'
$($CIM_OperatingSystem.Caption) | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</td></tr>'
#row end
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax">OS Version</td>
<td class="tg-0lax">'
$($CIM_OperatingSystem.Version) | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</td></tr>'
#row end    
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax">BuildNumber</td>
<td class="tg-0lax">'
$($CIM_OperatingSystem.BuildNumber) | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</td></tr>'
#row end
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax">ServicePack</td>
<td class="tg-0lax">'
$($CIM_OperatingSystem.ServicePackMajorVersion) | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</td></tr>'
#row end
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax">Uptime</td>
<td class="tg-0lax">'
$bootuptime = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime
$CurrentDate = Get-Date
$uptime = $CurrentDate - $bootuptime
"Days: $($uptime.days), Hours: $($uptime.Hours), Minutes: $($uptime.Minutes)" | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</td></tr>'
#row end
Add-Content -Path $FilePath -Value '</td></tr>'
#---SECTION END

#---SECTION START
Add-Content -Path $FilePath -Value '
	<tr>
	<th class="tg-0lax" colspan="2" id="Motherboard">Motherboard</th>
	</tr>
	</thead>
	<tbody>
	<tr>'

Add-Content -Path $FilePath -Value '<td class="tg-0lax">Motherboard Manufacturer:</td><td class="tg-0lax">'

$BaseBoardInfo.Manufacturer | Add-Content -Path $FilePath

Add-Content -Path $FilePath -Value '</td>
	</tr>
	<tr>
<td class="tg-0lax">Motherboard Model:</td>
<td class="tg-0lax">'
$BaseBoardInfo.Product | Add-Content -Path $FilePath

Add-Content -Path $FilePath -Value '</td></tr>'
Add-Content -Path $FilePath -Value '<tr>
	<td class="tg-0lax">BIOS/UEFI Ver.:</td>
	<td class="tg-0lax">'
Add-Content -Path $FilePath -Value $BIOSInfo.SMBIOSBIOSVersion
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax">WMI Motherboard:</td>
<td class="tg-0lax">'
Get-CimInstance -ClassName Win32_BaseBoard | Select-Object -Property Manufacturer, Model, PartNumber, SerialNumber, SKU, Version, Product | ConvertTo-Html -Fragment -As List | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</td></tr>'
#row end


Add-Content -Path $FilePath -Value '</td></tr>'
#---SECTION END
#---SECTION START
Add-Content -Path $FilePath -Value '
	<tr>
	<th class="tg-0lax" colspan="2" id="MEMORY">MEMORY</th>
	</tr>
	</thead>
	<tbody>
	<tr>'
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax">RAM</td>
<td class="tg-0lax">'
Get-CimInstance Win32_PhysicalMemoryArray | Select-Object -Property MaxCapacity,MemoryDevices | ConvertTo-Html -Fragment | Add-Content -Path $FilePath
Get-CimInstance win32_physicalmemory | ConvertTo-Html -Fragment -Property Manufacturer,Banklabel,Configuredclockspeed,Devicelocator,Capacity,Model,PartNumber,SerialNumber,MaxCapacity,MemoryDevices | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</td></tr>'
#row end

Add-Content -Path $FilePath -Value '</td></tr>'
#---SECTION END
#---SECTION START
Add-Content -Path $FilePath -Value '
	<tr>
	<th class="tg-0lax" colspan="2" id="HARDDRIVES">HARD DRIVES</th>
	</tr>
	</thead>
	<tbody>
	<tr>'
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax">Hard drives<br>Get-Disk</td>
<td class="tg-0lax">'
Get-Disk | ConvertTo-Html -Fragment -Property FriendlyName, Model, SerialNumber, HealthStatus, OperationalStatus, Size, TotalSize, IsSystem, PartitionStyle | Add-Content -Path $FilePath

$disks = Get-WmiObject -Class Win32_DiskDrive

$output = ""

foreach ($disk in $disks) {
    $output += "<h2>Disk: $($disk.DeviceID)</h2>"
    $output += "<table>"
    $output += "<tr><th>Property</th><th>Value</th></tr>"
    $output += "<tr><td><b>Model</b></td><td style='background-color:lightgreen;'>$($disk.Model)</td></tr>"
    $output += "<tr><td><b>Serial Number</b></td><td style='background-color:lightgreen;'>$($disk.SerialNumber)</td></tr>"
    $output += "<tr><th>Partition</th><th>Drive Letter</th><th>Size (GB)</th><th>Free Space (GB)</th><th>Used Space (GB)</th></tr>"

    $partitions = Get-WmiObject -Class Win32_DiskPartition | Where-Object {$_.DiskIndex -eq $disk.Index}

    foreach ($partition in $partitions) {
        $driveLetter = (Get-WmiObject -Query "ASSOCIATORS OF {Win32_DiskPartition.DeviceID='$($partition.DeviceID)'} WHERE AssocClass=Win32_LogicalDiskToPartition").DeviceID

        $sizeGB = [math]::Round($partition.Size / 1GB, 2)
        
        # Get free space from the associated logical disk
        $freeSpaceGB = [math]::Round((Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='$driveLetter'").FreeSpace / 1GB, 2)
        
        $usedSpaceGB = $sizeGB - $freeSpaceGB

        $output += "<tr>"
        $output += "<td>Partition $($partition.DeviceID)</td>"
        $output += "<td>$driveLetter</td>"
        $output += "<td>$sizeGB</td>"
        $output += "<td>$freeSpaceGB</td>"
        $output += "<td>$usedSpaceGB</td>"
        $output += "</tr>"
    }

    $output += "</table><br>"
}

$output | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</td></tr>'
#row end
Add-Content -Path $FilePath -Value '</td></tr>'
#---SECTION END
#---SECTION START
Add-Content -Path $FilePath -Value '
	<tr>
	<th class="tg-0lax" colspan="2" id="DISPLAY">DISPLAY</th>
	</tr>
	</thead>
	<tbody>
	<tr>'
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax">Display</td>
<td class="tg-0lax">'
Get-CimInstance win32_videocontroller | Select-Object -Property caption, VideoProcessor, CurrentHorizontalResolution, CurrentVerticalResolution | ConvertTo-Html -Fragment | Add-Content -Path $FilePath
Get-CimInstance win32_desktopmonitor | Select-Object -Property DeviceID, Name, ScreenHeight, ScreenWidth, Caption | ConvertTo-Html -Fragment | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</td></tr>'
#row end
Add-Content -Path $FilePath -Value '</td></tr>'
#---SECTION END
#---SECTION START
Add-Content -Path $FilePath -Value '
	<tr>
	<th class="tg-0lax" colspan="2" id="Network">NETWORK</th>
	</tr>
	</thead>
	<tbody>
	<tr>'
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax">Network</td>
<td class="tg-0lax">'
$ipconfig=ipconfig /all |%{"$_<br/>"}
#Add-Content -Path $FilePath -Value '<pre>'
$ipconfig | Add-Content -Path $FilePath
#Add-Content -Path $FilePath -Value '</pre>'
Add-Content -Path $FilePath -Value '</td></tr>'
#row end
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax" id="NetworkShares">Network Shares</td>
<td class="tg-0lax">'
get-smbshare | Select-Object -Property ShareState, Description, Name, Path, ShadowCopy | ConvertTo-Html -Fragment | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</td></tr>'
#row end
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax" id="Networkmappeddrives">Network mapped drives</td>
<td class="tg-0lax">'
Get-SMBMapping | select-object -Property Status, LocalPath, RemotePath, RequireIntegrity, RequirePrivacy, UseWriteThrough | ConvertTo-Html -Fragment | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</td></tr>'
#row end
Add-Content -Path $FilePath -Value '</td></tr>'
#---SECTION END
#---SECTION START
Add-Content -Path $FilePath -Value '
	<tr>
	<th class="tg-0lax" colspan="2">PRINTERS</th>
	</tr>
	</thead>
	<tbody>
	<tr>'
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax" id="PRINTERS">Printers</td>
<td class="tg-0lax">'
Get-Printer | Select-Object -Property Type, DeviceType, Caption, Description, Name, PrimaryStatus, Comment, Datatype, DriverName, KeepPrintedJobs, PortName, PrintProcessor, Shared, ShareName | ConvertTo-Html -Fragment -As Table | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</td></tr>'
#row end
Add-Content -Path $FilePath -Value '</td></tr>'
#---SECTION END
#---SECTION START
Add-Content -Path $FilePath -Value '
	<tr>
	<th class="tg-0lax" colspan="2" id="SOFTWARE">SOFTWARE</th>
	</tr>
	</thead>
	<tbody>
	<tr>'
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax">Installed Software</td>
<td class="tg-0lax">'
$computername = $env:COMPUTERNAME
$uninstallkey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
$reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine', $computername)
$regkey = $reg.OpenSubKey($uninstallkey)
$subkeys = $regkey.GetSubKeyNames()
foreach($key in $subkeys) {
	$thiskey = $uninstallkey + "\\" + $key
	$thissubkey = $reg.OpenSubKey($thiskey)
	$DisplayName = $thissubkey.GetValue("DisplayName")
	$DisplayName + "<br>" | Add-Content -Path $FilePath
}
Add-Content -Path $FilePath -Value '</td></tr>'
#row end
Add-Content -Path $FilePath -Value '</td></tr>'
#---SECTION END

#---SECTION START
Add-Content -Path $FilePath -Value '
	<tr>
	<th class="tg-0lax" colspan="2">SKAR_specs (C) 2024</th>
	</tr>
	</thead>
	<tbody>
	<tr>'
#row start
Add-Content -Path $FilePath -Value '<tr>
<td class="tg-0lax">Copyright</td>
<td class="tg-0lax">Gather information about computer hardware and software. Output to a HTML file.</td></tr>'
#row end
Add-Content -Path $FilePath -Value '</td></tr>'
#---SECTION END

### END
Add-Content -Path $FilePath -Value '</tbody></table></div></body></html>'
### END

if (test-path $FilePath) {
	Get-ChildItem -Path $FilePath | Rename-Item -NewName { [System.IO.Path]::ChangeExtension($_.Name, ".html") }
}