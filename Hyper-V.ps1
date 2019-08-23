Clear-Host
# Get information about the current culture settings
# Получить сведения о параметрах текущей культуры
IF ((Get-Culture).Name -eq "ru-RU")
{
	$RU = $true
}
# Enable Hyper-V
# Включить Hyper-V
IF ((Get-CimInstance –ClassName CIM_ComputerSystem).HypervisorPresent -eq $false)
{
	Enable-WindowsOptionalFeature -FeatureName Microsoft-Hyper-V-All -Online -NoRestart
	IF ($RU)
	{
		Write-Host "`nПерезагрузите ПК"
	}
	else
	{
		Write-Host "`nRestart the PC"
	}
	break
}
IF ((Get-CimInstance –ClassName CIM_ComputerSystem).HypervisorPresent -eq $true)
{
	IF ($RU)
	{
		Write-Host "`nВведите имя для ВМ"
		Write-Host "Команда `"list`" выводит список имеющихся ВМ" -NoNewline
	}
	else
	{
		Write-Host "`nType name for a VM"
		Write-Host "`"list`" command displays VM names list" -NoNewline
	}
	$VMName = Read-Host -Prompt " "
	IF ($VMName -eq "list")
	{
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
		break
	}
	$VirtualHardDiskPath = (Get-VMHost).VirtualHardDiskPath
	IF ((Get-VM -VMName $VMName -ErrorAction SilentlyContinue) -or (Test-Path -Path $VirtualHardDiskPath\$VMName))
	{
		IF ($RU)
		{
			Write-Host "`nВМ `"$VMName`" уже существует."
			Write-Host "Удалить ВМ `"$VMName`" и папку ВМ $VirtualHardDiskPath\$VMName`?"
			Write-Host "Введите `"yes`", чтобы удалить" -ForegroundColor Yellow
			Write-Host "`nЧтобы пропустить, нажмите Enter" -NoNewline
		}
		else
		{
			Write-Host "`nVM `"$VMName`" already exists."
			Write-Host "Delete VM `"$VMName`" and VM folder $VirtualHardDiskPath\$VMName`?"
			Write-Host "Type `"yes`" to delete" -ForegroundColor Yellow
			Write-Host "`nPress Enter to skip" -NoNewline
		}
		$command = Read-Host -Prompt " "
		IF ($command -eq "yes")
		{
			Get-VM -VMName $VMName -ErrorAction SilentlyContinue | Where-Object -FilterScript {$_.State -eq "Running"} | Stop-VM -Force
			Remove-VM -VMName $VMName -Force -ErrorAction SilentlyContinue
			Remove-Item -Path "$VirtualHardDiskPath\$VMName" -Recurse -Force -ErrorAction SilentlyContinue
		}
		elseif ([string]::IsNullOrEmpty($command))
		{
			break
		}
		else
		{
			Write-Host "`nInvalid command" -ForegroundColor Yellow
			break
		}
	}
	# Set default location for virtual hard disk to "$env:SystemDrive\HV"
	# Установить папку по умолчанию для виртуальных жестких дисков на "$env:SystemDrive\HV"
	$VirtualHardDiskPath = "$env:SystemDrive\HV"
	IF (-not (Test-Path $VirtualHardDiskPath))
	{
		New-Item -ItemType Directory -Path $VirtualHardDiskPath
	}
	IF ((Get-VMHost).VirtualHardDiskPath -ne $VirtualHardDiskPath)
	{
		Set-VMHost -VirtualHardDiskPath $VirtualHardDiskPath
	}
	# Set default location for virtual VMs to "$env:SystemDrive\HV"
	# Установить папку по умолчанию для виртуальных машин
	IF ((Get-VMHost).VirtualMachinePath -ne $VirtualHardDiskPath)
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
	# Добавить .iso образ ОС
	Add-Type -AssemblyName System.Windows.Forms
	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	# Downloads folder
	# Папка "Загрузки"
	$DownloadsFolder = Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" -Name "{374DE290-123F-4565-9164-39C4925E467B}"
	$OpenFileDialog.InitialDirectory = $DownloadsFolder
	$OpenFileDialog.Multiselect = $false
	IF ($RU)
	{
		$OpenFileDialog.Filter = "ISO файлы (*.iso)|*.iso|Все файлы (*.*)|*.*"
	}
	else
	{
		$OpenFileDialog.Filter = "ISO Files (*.iso)|*.iso|All Files (*.*)|*.*"
	}
	$OpenFileDialog.ShowHelp = $true
	$OpenFileDialog.ShowDialog()
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
	# Set the amount of RAM for VM half as much as installed
	# Set the number of virtual processors for VM to $env:NUMBER_OF_PROCESSORS
	Set-VMProcessor -VMName $VMName -Count $env:NUMBER_OF_PROCESSORS
	# Create virtual switch
	# Создать внешний виртуальный коммутатор
	IF ((Get-VMSwitch -SwitchType External).NetAdapterInterfaceDescription -ne (Get-NetAdapter -Physical).InterfaceDescription)
	{
		New-VMSwitch -Name "Virtual switch" -NetAdapterName (Get-NetAdapter -Physical).Name -AllowManagementOS $true
	}
	# Set virtual switch for VM
	# Установить виртуальный коммутатор для ВМ
	Get-VM -VMName $VMName | Get-VMNetworkAdapter | Connect-VMNetworkAdapter -SwitchName (Get-VMSwitch -SwitchType External).Name
	# Do not use automatic checkpoints
	# Не использовать автоматические контрольные точки для ВМ
	Set-VM -VMName $VMName -AutomaticCheckpointsEnabled $false
	# Set the initial VM boot from DVD drive
	# Установить первоначальную загрузку ВМ с DVD-дисковода
	Set-VMFirmware -VMName $VMName -FirstBootDevice $(Get-VMDvdDrive -VMName $VMName)
	# Enable nested virtualization for VM
	# Разрешить вложенную виртуализацию
	Set-VMProcessor -VMName $VMName -ExposeVirtualizationExtensions $true
}
# Connect to VM
# Подключиться к ВМ
vmconnect.exe $env:COMPUTERNAME $VMName
# Start VM
# Запустить ВМ
Start-Sleep -Seconds 5
Start-VM -VMName $VMName
# Show vmconnect.exe window to send space key
# Вывести на передний план окно vmconnect.exe, чтобы послать нажатие виртуального пробела
$Win32ShowWindowAsync = @{
	Namespace = "Win32Functions"
	Name = "Win32ShowWindowAsync"
	Language = "CSharp"
	MemberDefinition = @"
		[DllImport("user32.dll")]
		public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
		[DllImport("user32.dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool SetForegroundWindow(IntPtr hWnd);
"@
}
IF (-not ("Win32Functions.Win32ShowWindowAsync" -as [type]))
{
	Add-Type @Win32ShowWindowAsync
}
$title = "*$VMName*$env:COMPUTERNAME*"
Get-Process | ForEach-Object -Process {
	IF (($_.ProcessName -eq "vmconnect") -and ($_.MainWindowTitle -like $title))
	{
		# Show window, if minimized
		# Развернуть окно, если свернуто
		[Win32Functions.Win32ShowWindowAsync]::ShowWindowAsync($_.MainWindowHandle, 10) | Out-Null
		# Move focus to the window
		# Перевести фокус на окно
		[Win32Functions.Win32ShowWindowAsync]::SetForegroundWindow($_.MainWindowHandle) | Out-Null
	}
}
# Emulate space key sending to initialize OS installing
# Эмулировать нажатие виртуального пробела, чтобы инициализировать установку
Start-Sleep -Seconds 1
[System.Windows.Forms.SendKeys]::SendWait(" ")
# Edit session settings
# Изменить настройки сессии
# vmconnect.exe $env:COMPUTERNAME $VMName /edit

# Expand HDD space to 40 GB after OS installed
# Расширить объем ж/д до 40 ГБ после установки ОС
# (Get-VM -VMName $VMName).HardDrives | Select-Object -First 1 | Resize-VHD -SizeBytes 40gb -Passthru
