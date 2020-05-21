Clear-Host

#region Begin
# Enable Hyper-V
# Включить Hyper-V
# Check whether Hyper-V is enabled
if ((Get-WindowsOptionalFeature -FeatureName Microsoft-Hyper-V-All -Online).State -eq "Disabled")
{
	Enable-WindowsOptionalFeature -FeatureName Microsoft-Hyper-V-All -Online -NoRestart
	Write-Warning -Message "Restart the PC"
	break
}
#endregion Begin

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

if ((Get-VM -VMName $VMName -ErrorAction SilentlyContinue) -or (Test-Path -Path $VirtualHardDiskPath\$VMName))
{
	Write-Output "`nVM `"$VMName`" already exists."
	Write-Output "Delete VM `"$VMName`" and VM folder $VirtualHardDiskPath\$VMName`?"
	$yes = '"{0}"' -f "yes"
	Write-Output "`nPress Enter to skip"
	$command = Read-Host -Prompt "Type $yes to delete"
	if ([string]::IsNullOrEmpty($command))
	{
		break
	}
	else
	{
		switch ($command)
		{
			"yes"
			{
				Get-VM -VMName $VMName -ErrorAction SilentlyContinue | Where-Object -FilterScript {$_.State -eq "Running"} | Stop-VM -Force
				Remove-VM -VMName $VMName -Force -ErrorAction SilentlyContinue
				Remove-Item -Path "$VirtualHardDiskPath\$VMName" -Recurse -Force -ErrorAction SilentlyContinue
			}
			Default
			{
				Write-Warning -Message "Invalid command"
				return
			}
		}
	}
}
#endregion VMName

#region Settings
# Set default location for virtual hard disk to "$env:SystemDrive\HV"
# Установить папку по умолчанию для виртуальных жестких дисков на "$env:SystemDrive\HV"
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
# Установить папку по умолчанию для виртуальных машин
if ((Get-VMHost).VirtualMachinePath -ne $VirtualHardDiskPath)
{
	Set-VMHost -VirtualMachinePath $VirtualHardDiskPath
}

# Create gen 2 VM
# Создать виртуальную машину второго поколения
New-VM -VMName $VMName -Path $VirtualHardDiskPath\$VMName -Generation 2

# Create a 30 GB virtual hard drive
# Создать виртуальный жесткий диск размером 30 ГБ
New-VHD -Dynamic -SizeBytes 30GB -Path "$VirtualHardDiskPath\$VMName\VirtualHardDisk\$VMName.vhdx"

# Add a hard disk drive to a virtual machine
# Присоединить виртуальный жесткий диск к виртуальной машине
Add-VMHardDiskDrive -VMName $VMName -Path "$VirtualHardDiskPath\$VMName\VirtualHardDisk\$VMName.vhdx"

# Add .iso image
# Добавить .iso образ
Add-Type -AssemblyName System.Windows.Forms
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
# Downloads folder
# Папка "Загрузки"
$DownloadsFolder = Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" -Name "{374DE290-123F-4565-9164-39C4925E467B}"
$OpenFileDialog.InitialDirectory = $DownloadsFolder
$OpenFileDialog.Filter = "ISO Files (*.iso)|*.iso|All Files (*.*)|*.*"
# Focus on open file dialog
# Перевести фокус на диалог открытия файла
$tmp = New-Object System.Windows.Forms.Form -Property @{TopMost = $true}
$OpenFileDialog.ShowDialog($tmp)

if ($OpenFileDialog.FileName)
{
	# Use selected .iso image with virtual CD\DVD drive
	# Использовать выбранный .iso образ для виртуального дисковода компакт- и DVD-дисков
	Add-VMDvdDrive -VMName $VMName -Path $OpenFileDialog.FileName

	# Get localized name of "Guest Service Interface"
	# Получить локализованное имя "Интерфейса гостевой службы"
	$guestServiceId = "Microsoft:{0}\6C09BB55-D683-4DA0-8931-C9BF705F6480" -f (Get-VM -VMName $VMName).Id
	$Name = (Get-VMIntegrationService -VMName $VMName | Where-Object -FilterScript {$_.Id -eq $guestServiceId}).Name

	# Enable "Guest Service Interface" for VM
	# Включить "Интерфейс гостевой службы" для ВМ
	Get-VM -VMName $VMName | Enable-VMIntegrationService -Name $Name

	# Set the amount of RAM for VM half as much as installed
	# Установить объем оперативной памяти для ВМ вдвое меньше, чем установлено
	$ram = ((Get-CimInstance -ClassName CIM_PhysicalMemory).Capacity | Measure-Object -Sum).Sum
	Set-VMMemory -VMName $VMName -StartupBytes $($ram/2)

	# Set the number of virtual processors for VM to $env:NUMBER_OF_PROCESSORS
	# Установить число виртуальных прцоессоров на $env:NUMBER_OF_PROCESSORS
	Set-VMProcessor -VMName $VMName -Count $env:NUMBER_OF_PROCESSORS

	# Create external virtual switch
	# Создать внешний виртуальный коммутатор
	if ((Get-VMSwitch -SwitchType External).NetAdapterInterfaceDescription -ne (Get-NetAdapter -Physical).InterfaceDescription)
	{
		New-VMSwitch -Name "Virtual switch" -NetAdapterName (Get-NetAdapter -Physical).Name -AllowManagementOS $true
	}

	# Set virtual switch for VM
	# Установить виртуальный коммутатор для ВМ
	Get-VM -VMName $VMName | Get-VMNetworkAdapter | Connect-VMNetworkAdapter -SwitchName (Get-VMSwitch -SwitchType External).Name

	# Do not use automatic checkpoints for VM
	# Не использовать автоматические контрольные точки для ВМ
	Set-VM -VMName $VMName -AutomaticCheckpointsEnabled $false

	# Verifying .iso image
	# Проверка .iso-образа
	$Title = "Graphics performance preference"
	$Message = "Original Microsoft .iso image?"
	$Options = "&Yes", "&No"
	$DefaultChoice = 1
	$Result = $Host.UI.PromptForChoice($Title, $Message, $Options, $DefaultChoice)
	switch ($Result)
	{
		"0"
		{
			# Original .iso
			Set-VMFirmware -VMName $VMName -EnableSecureBoot On
		}
		"1"
		{
			# Custom compiled .iso
			Set-VMFirmware -VMName $VMName -EnableSecureBoot Off
		}
	}

	# Set the initial VM boot from DVD drive
	# Установить первоначальную загрузку ВМ с DVD-дисковода
	Set-VMFirmware -VMName $VMName -FirstBootDevice $(Get-VMDvdDrive -VMName $VMName)

	# Set boot order: Dvd Drive, Hard Disk, Network Adapter
	# Установить порядок загрузки: Dvd Drive, Hard Disk, Network Adapter
	$VMDvdDrive = Get-VMDvdDrive -VMName $VMName
	$VMHardDiskDrive = Get-VMHardDiskDrive -VMName $VMName
	$VMNetworkAdapter = Get-VMNetworkAdapter -VMName $VMName
	Set-VMFirmware -VMName $VMName -BootOrder $VMDvdDrive, $VMHardDiskDrive, $VMNetworkAdapter

	# Enable nested virtualization for VM
	# Разрешить вложенную виртуализацию
	Set-VMProcessor -VMName $VMName -ExposeVirtualizationExtensions $true

	# Connect to VM
	# Подключиться к ВМ
	vmconnect.exe $env:COMPUTERNAME $VMName

	# Start VM
	# Запустить ВМ
	Start-Sleep -Seconds 5
	Start-VM -VMName $VMName

	#region Window
	# Set vmconnect.exe window to the foreground to send space key
	# Вывести на передний план окно vmconnect.exe, чтобы послать нажатие виртуального пробела
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
	$title = "*$VMName*$env:COMPUTERNAME*"
	Get-Process | ForEach-Object -Process {
		if (($_.ProcessName -eq "vmconnect") -and ($_.MainWindowTitle -like $title))
		{
			# Show window, if minimized
			# Развернуть окно, если свернуто
			[WinAPI.ForegroundWindow]::ShowWindowAsync($_.MainWindowHandle, 10)
			Start-Sleep -Milliseconds 100

			# Move focus to the window
			# Перевести фокус на окно
			[WinAPI.ForegroundWindow]::SetForegroundWindow($_.MainWindowHandle)

			# Emulate Enter key sending 100 times to initialize OS installing
			# Эмулировать нажатие Enter 100 раз, чтобы инициализировать установку
			Start-Sleep -Milliseconds 100
			[System.Windows.Forms.SendKeys]::SendWait("{Enter 100}")
		}
	}
	#endregion Window
}
#endregion VMName

# Edit session settings
# Изменить настройки сессии
# vmconnect.exe $env:COMPUTERNAME $VMName /edit

# Expand HDD space to 40 GB after OS installed
# Расширить объем ж/д до 40 ГБ после установки ОС
# (Get-VM -VMName $VMName).HardDrives | Select-Object -First 1 | Resize-VHD -SizeBytes 40gb -Passthru