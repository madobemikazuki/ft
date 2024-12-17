Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

. ..\ft_cores\FT_IO.ps1
. ..\ft_cores\Watcher\FT_Ezy_Watcher.ps1

$script:config = [FT_IO]::Read_JSON_Object(".\config\ezy_watch_commander.json")

$script:log_file_path = (${HOME} + $config.log_file)
#Test-Path $log_file_path

$script:target_objects = $config.watch_targets


$script:gZEN_path = (${HOME} + $target_objects.rsv.path)
$script:gZEN_orders = $target_objects.rsv.orders
$script:gZEN_block = {
  foreach ($_order in $GZEN_orders){
    . $_order
  }
}

$script:order = @{
  path = $gZEN_path;
  orders = $gZEN_orders;
  action_block = $gZEN_Block;
}
Set-Location ($target_objects.rsv).current_directory

[FT_Ezy_Watcher]::Start(@($order))

