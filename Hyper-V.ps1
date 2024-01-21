#Requires -RunAsAdministrator

Clear-Host

# Check whether Hyper-V is enabled
if ((Get-WindowsOptionalFeature -FeatureName Microsoft-Hyper-V-All -Online).State -eq "Disabled")
{
	Write-Warning -Message "Restart the PC"

	break
}

#region VMName
Write-Output "Available VMs"

$Name = @{
	Name = "VM Name"
	Expression = {$_.Name}
}
$Path = @{
	Name = "Path"
	Expression = {$_.Path}
}
$State = @{
	Name = "VM State"
	Expression = {$_.State}
}
(Get-VM | Select-Object -Property $Name, $Path, $State | Format-Table | Out-String).Trim()

$VMName = Read-Host -Prompt "`nType name for a VM"

$VirtualHardDiskPath = (Get-VMHost).VirtualHardDiskPath

#region Show-Menu
function script:Show-Menu
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[array]
		$Menu,

		[Parameter(Mandatory = $true)]
		[int]
		$Default
	)

	Write-Information -MessageData "" -InformationAction Continue

	# Add "Please use the arrow keys 🠕 and 🠗 on your keyboard to select your answer" to menu
	$Menu += "Please use the arrow keys {0} and {1} on your keyboard to select your answer" -f [System.Char]::ConvertFromUtf32(0x2191), [System.Char]::ConvertFromUtf32(0x2193)

	# https://github.com/microsoft/terminal/issues/14992
	[System.Console]::BufferHeight += $Menu.Count
	$minY = [Console]::CursorTop
	$y = [Math]::Max([Math]::Min($Default, $Menu.Count), 0)
	do
	{
		[Console]::CursorTop = $minY
		[Console]::CursorLeft = 0
		$i = 0
		foreach ($item in $Menu)
		{
			if ($i -ne $y)
			{
				Write-Information -MessageData ('  {1}  ' -f ($i+1), $item) -InformationAction Continue
			}
			else
			{
				Write-Information -MessageData ('[ {1} ]' -f ($i+1), $item) -InformationAction Continue
			}
			$i++
		}
		$k = [Console]::ReadKey()
		switch ($k.Key)
		{
			"UpArrow"
			{
				if ($y -gt 0)
				{
					$y--
				}
			}
			"DownArrow"
			{
				if ($y -lt ($Menu.Count - 1))
				{
					$y++
				}
			}
			"Enter"
			{
				return $Menu[$y]
			}
		}
	}
	while ($k.Key -notin ([ConsoleKey]::Escape, [ConsoleKey]::Enter))
}
#endregion Show-Menu

if ((Get-VM -VMName $VMName -ErrorAction Ignore) -or (Test-Path -Path $VirtualHardDiskPath\$VMName))
{
	Write-Verbose "VM `"$VMName`" already exists" -Verbose

	Write-Information -MessageData "" -InformationAction Continue
	Write-Verbose -Message ("Delete VM `"$VMName`" and VM folder $VirtualHardDiskPath\$VMName`?") -Verbose

	$Script:KeyboardArrows = "Please use the arrow keys {0} and {1} on your keyboard to select your answer" -f [System.Char]::ConvertFromUtf32(0x2191), [System.Char]::ConvertFromUtf32(0x2193)
	$Script:Yes = "Yes"
	$Script:Skip = "Skip"

	do
	{
		$Choice = Show-Menu -Menu @($Yes, $Skip) -Default 2

		switch ($Choice)
		{
			$Yes
			{
				Get-VM -VMName $VMName -ErrorAction Ignore | Where-Object -FilterScript {$_.State -eq "Running"} | Stop-VM -Force
				Remove-VM -VMName $VMName -Force -ErrorAction Ignore
				Remove-Item -Path "$VirtualHardDiskPath\$VMName" -Recurse -Force -ErrorAction Ignore
			}
			$Skip
			{
				Write-Verbose "Skipped" -Verbose

				return
			}
			$KeyboardArrows {}
		}
	}
	until ($Choice -ne $KeyboardArrows)
}
#endregion VMName

#region Settings
# Set default location for virtual hard disk to "$env:SystemDrive\HV"
$VirtualHardDiskPath = "$env:SystemDrive\HV"
if (-not (Test-Path $VirtualHardDiskPath))
{
	New-Item -ItemType Directory -Path $VirtualHardDiskPath
}
if ((Get-VMHost).VirtualHardDiskPath -ne $VirtualHardDiskPath)
{
	Set-VMHost -VirtualHardDiskPath $VirtualHardDiskPath
}

# Set default location for virtual VMs to "$env:SystemDrive\HV"
if ((Get-VMHost).VirtualMachinePath -ne $VirtualHardDiskPath)
{
	Set-VMHost -VirtualMachinePath $VirtualHardDiskPath
}

# Create a gen 2 VM
New-VM -VMName $VMName -Path $VirtualHardDiskPath\$VMName -Generation 2

# Enable vTPM
if (-not (Get-HgsGuardian -Name UntrustedGuardian -ErrorAction Ignore))
{
    # Creating new UntrustedGuardian since it did not exist
    New-HgsGuardian -Name UntrustedGuardian -GenerateCertificates
}
$RawData = (New-HgsKeyProtector -Owner (Get-HgsGuardian -Name UntrustedGuardian -ErrorAction Stop) -AllowUntrustedRoot).RawData 
Set-VMKeyProtector -VMName $VMName -KeyProtector $RawData
Enable-VMTPM -VMName $VMName

# Create a 40 GB virtual hard drive
New-VHD -Dynamic -SizeBytes 52GB -Path "$VirtualHardDiskPath\$VMName\VirtualHardDisk\$VMName.vhdx"

# Add a hard disk drive to a virtual machine
Add-VMHardDiskDrive -VMName $VMName -Path "$VirtualHardDiskPath\$VMName\VirtualHardDisk\$VMName.vhdx"

# Add an .iso image
Add-Type -AssemblyName System.Windows.Forms
$OpenFileDialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
$DownloadsFolder = Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" -Name "{374DE290-123F-4565-9164-39C4925E467B}"
$OpenFileDialog.InitialDirectory = $DownloadsFolder
$OpenFileDialog.Filter = "ISO Files (*.iso)|*.iso|All Files (*.*)|*.*"
# Focus on open file dialog
$tmp = New-Object System.Windows.Forms.Form -Property @{TopMost = $true}
$OpenFileDialog.ShowDialog($tmp)

if ($OpenFileDialog.FileName)
{
	# Use selected .iso image with virtual CD\DVD drive
	Add-VMDvdDrive -VMName $VMName -Path $OpenFileDialog.FileName

	# Get localized name of "Guest Service Interface"
	$guestServiceId = "Microsoft:{0}\6C09BB55-D683-4DA0-8931-C9BF705F6480" -f (Get-VM -VMName $VMName).Id
	$Name = (Get-VMIntegrationService -VMName $VMName | Where-Object -FilterScript {$_.Id -eq $guestServiceId}).Name

	# Enable "Guest Service Interface" for VM
	Get-VM -VMName $VMName | Enable-VMIntegrationService -Name $Name

	# Set the amount of RAM for VM half as much as installed
	$RAM = ((Get-CimInstance -ClassName CIM_PhysicalMemory).Capacity | Measure-Object -Sum).Sum
	Set-VMMemory -VMName $VMName -StartupBytes $($RAM/4)

	# Set the number of virtual processors for VM to $env:NUMBER_OF_PROCESSORS
	Set-VMProcessor -VMName $VMName -Count $($env:NUMBER_OF_PROCESSORS/4)

	# Create an external virtual switch
	if (-not (Get-VMSwitch -SwitchType External -Name "Virtual switch" -ErrorAction Ignore))
	{
		$WiredInterface = Get-NetAdapter -Physical | Where-Object {$_.PhysicalMediaType -eq "802.3"}
		New-VMSwitch -Name "Virtual switch" -NetAdapterName $WiredInterface.Name -AllowManagementOS $true
	}

	# Set virtual switch for VM
	Get-VM -VMName $VMName | Get-VMNetworkAdapter | Connect-VMNetworkAdapter -SwitchName (Get-VMSwitch -SwitchType External).Name

	# Do not use automatic checkpoints for VM
	Set-VM -VMName $VMName -AutomaticCheckpointsEnabled $false

	# Verifying .iso image
	Write-Information -MessageData "" -InformationAction Continue
	Write-Verbose -Message ("Original Microsoft .iso image?") -Verbose

	$Script:KeyboardArrows = "Please use the arrow keys {0} and {1} on your keyboard to select your answer" -f [System.Char]::ConvertFromUtf32(0x2191), [System.Char]::ConvertFromUtf32(0x2193)
	$Script:Yes = "Yes"
	$Script:No = "No"

	do
	{
		$Choice = Show-Menu -Menu @($Yes, $No) -Default 2

		switch ($Choice)
		{
			$Yes
			{
				# Original .iso
				Set-VMFirmware -VMName $VMName -EnableSecureBoot On
			}
			$No
			{
				# Custom compiled .iso
				Set-VMFirmware -VMName $VMName -EnableSecureBoot Off
			}
			$KeyboardArrows {}
		}
	}
	until ($Choice -ne $KeyboardArrows)

	# Set the initial VM boot from DVD drive
	Set-VMFirmware -VMName $VMName -FirstBootDevice $(Get-VMDvdDrive -VMName $VMName)

	# Set boot order: DVD Drive, Hard Disk, Network Adapter
	$VMDvdDrive = Get-VMDvdDrive -VMName $VMName
	$VMHardDiskDrive = Get-VMHardDiskDrive -VMName $VMName
	$VMNetworkAdapter = Get-VMNetworkAdapter -VMName $VMName
	Set-VMFirmware -VMName $VMName -BootOrder $VMDvdDrive, $VMHardDiskDrive, $VMNetworkAdapter

	# Enable nested virtualization for VM
	Set-VMProcessor -VMName $VMName -ExposeVirtualizationExtensions $true

	# Connect to VM
	vmconnect.exe $env:COMPUTERNAME $VMName

	# Start VM
	Start-Sleep -Seconds 5
	Start-VM -VMName $VMName

	#region Window
	# Set vmconnect.exe window to the foreground to send space key
	$SetForegroundWindow = @{
		Namespace = "WinAPI"
		Name = "ForegroundWindow"
		Language = "CSharp"
		MemberDefinition = @"
			[DllImport("user32.dll")]
			public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
			[DllImport("user32.dll")]
			[return: MarshalAs(UnmanagedType.Bool)]
			public static extern bool SetForegroundWindow(IntPtr hWnd);
"@
	}
	if (-not ("WinAPI.ForegroundWindow" -as [type]))
	{
		Add-Type @SetForegroundWindow
	}

	Get-Process | Where-Object -FilterScript {($_.ProcessName -eq "vmconnect") -and ($_.MainWindowTitle -like "*$VMName*$env:COMPUTERNAME*")} | ForEach-Object -Process {
		# Show window, if minimized
		[WinAPI.ForegroundWindow]::ShowWindowAsync($_.MainWindowHandle, 10)

		Start-Sleep -Seconds 1

		# Force move the console window to the foreground
		[WinAPI.ForegroundWindow]::SetForegroundWindow($_.MainWindowHandle)

		Start-Sleep -Seconds 1

		# Emulate the Enter key sending 100 times to initialize OS installing
		[System.Windows.Forms.SendKeys]::SendWait("{Enter 100}")
	}
	#endregion Window
}
#endregion VMName

# Edit session settings
# vmconnect.exe $env:COMPUTERNAME $VMName /edit

# Expand HDD space to XX GB after OS installed
# (Get-VM -VMName $VMName).HardDrives | Select-Object -First 1 | Resize-VHD -SizeBytes XXgb -Passthru
