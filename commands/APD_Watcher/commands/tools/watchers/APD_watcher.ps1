Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [ValidatePattern("^[0-9]{2}\-[0-9]{6}$")]
  [String[]]$_central_nums
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

. ..\..\ft_cores\FT_IO.ps1
. ..\..\ft_cores\FT_Path.ps1
. ..\..\ft_cores\FT_Message.ps1

# 前回起動時のEventSubscriberを全て消す
Get-EventSubscriber -Force | Unregister-Event -Force


$private:command_name = Split-Path -Leaf $PSCommandPath
[FT_Message]::execution($command_name)

$script:APD_Config = [FT_IO]::Read_JSON_Object(".\config\csvAPD.json")


$script:APD_watcher = New-Object System.IO.FileSystemWatcher
$APD_watcher.Path = [FT_Path]::Fixed_Path($APD_Config.folder)
$APD_watcher.Filter = $APD_Config.file_name
$APD_watcher.IncludeSubdirectories = $false
$APD_watcher.EnableRaisingEvents = $True

$script:watcher_name = ("Created_" + $command_name)

$action = {
  Set-Location ..
  . .\csvAPD.ps1 $Event.MessageData
  Set-Location .\watchers
}

$reg_dict = @{
  InputObject      = $APD_watcher
  EventName        = "Created"
  SourceIdentifier = $watcher_name
  Action           = $action
  MessageData      = $_central_nums
}

Register-ObjectEvent @reg_dict
Write-Host $watcher_name" 起動"
while ($true) {
  Start-Sleep 1
}

