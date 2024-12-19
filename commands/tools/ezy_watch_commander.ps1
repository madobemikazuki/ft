Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

. ..\ft_cores\FT_IO.ps1
. ..\ft_cores\Watcher\FT_Ezy_Watcher.ps1

$private:config = [FT_IO]::Read_JSON_Object(".\config\ezy_watch_commander.json")


$script:log_file_path = (${HOME} + $config.log_file)
#Test-Path $log_file_path

$script:target_objects = $config.watch_targets

$contains_anpersand = {
  $_.Contains("&") -or $_.Contains("&")
}
$contains_Downloads = {
  $_.Contains("Downloads")
}

#$script:gZEN_path = (${HOME} + $target_objects.rsv.path)
$path = $target_objects.rsv.path
$script:gZEN_path = switch ($path) {
  { & $contains_anpersand } { $_ -replace "&|＆", "$"; break; }
  { & $contains_Downloads } { (${HOME} + $_); break; }
  default { $_ }
}


$script:gZEN_orders = $target_objects.rsv.orders
$script:gZEN_block = {
  . .\dl.ps1
  Set-Location ($target_objects.rsv).current_directory
  foreach ($_order in $GZEN_orders) {
    . $_order
  }
}

$script:order = @{
  path         = $gZEN_path;
  orders       = $gZEN_orders;
  action_block = $gZEN_Block;
}


[FT_Ezy_Watcher]::Start(@($order))

