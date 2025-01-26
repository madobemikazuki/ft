Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

. ..\ft_cores\FT_IO.ps1
$private:config = [FT_IO]::Read_JSON_Object(".\config\pcinfo.json")
$bgcolor = "DarkBlue"

Write-Host "COMPUTER 情報" -NoNewline -BackgroundColor $bgcolor
Get-WmiObject -Class Win32_ComputerSystem | Format-List $config.COMPUTER_INFO

Write-Host "CPU 情報" -NoNewline -BackgroundColor $bgcolor
Get-WmiObject Win32_Processor | Format-List -Property $config.CPU_INFO

Write-Host "OS 情報" -NoNewline -BackgroundColor $bgcolor
Get-WmiObject -Class Win32_OperatingSystem | Format-List -Property $config.OS_INFO

Write-Host "MEMORY 情報" -NoNewline -BackgroundColor $bgcolor
Get-WmiObject Win32_PhysicalMemory | Format-List -Property $config.MEMORY_INFO

Write-Host "BIOS 情報" -NoNewline -BackgroundColor $bgcolor
Get-WmiObject Win32_BIOS | Format-List -Property $config.BIOS_INFO

Remove-Variable config

