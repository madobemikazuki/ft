Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

. ..\ft_cores\FT_IO.ps1
. ..\ft_cores\FT_Message.ps1
. ..\ft_cores\Watcher\FT_Ezy_Watcher.ps1


function fn_Fixed_Path{
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$_folder_path
  )
  $private:parsed_folder_path = switch ($_folder_path) {
    #{ & $contains_anpersand } { $_ -replace "＆", "$"; break; }
    # 実行環境ではこのコードでは動かない？
    { $_.Contains("&") -or $_.Contains("＆") } { $_ -replace "&|＆", "$"; break; }
    { $_.Contains("Downloads") } { (${HOME} + $_); break; }
    default { $_ }
  }
  return $parsed_folder_path
}


# TODO: 実装中
function fn_CopyToDL {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$_from_path,
    [Parameter(Mandatory = $True, Position = 1)]
    [String]$_destination_path
  )
  Write-Host "　From: " $_from_path
  $private:file_name = Split-Path $_from_path -Leaf 
  $private:dist_path = ($_destination_path + $file_name)  
  Write-Host "　　To: " $dist_path
  Copy-Item -Path $_from_path -Destination $dist_path -Force
}


$private:config = [FT_IO]::Read_JSON_Object(".\config\ezy_watch_commander.json")
$script:log_file_path = (${HOME} + $config.log_file)
#Test-Path $log_file_path


$script:target_objects = $config.watch_targets.rsv
#$script:rsv_folder = (${HOME} + $target_objects.rsv.path)
# FIXME: 動作確認中
$private:folder_path = $target_objects.target_folder
$private:rsv_file = $target_objects.target_file_name
#Write-Host $folder_path
$script:rsv_folder = fn_Fixed_Path $folder_path

$rsv_path = [FT_IO]::Find_One($rsv_folder, $rsv_file)
#$rsv_path
$script:gZEN_orders = $target_objects.orders

$destination = fn_Fixed_Path $target_objects.destination

$script:gZEN_block = {
  # TODO: 監視対象のファイルをダウンロードする。
  #. .\dl.ps1

  fn_CopyToDL $rsv_path $destination


  Set-Location $target_objects.commands_directory
  foreach ($_order in $GZEN_orders) {
    . $_order
  }
  # FIXME: 文字化けするのでエンコードについて考慮すること
  #Write-OutPut "$(Get-Date), $rsv_path" >> $log_file_path
  Set-Location $PSScriptRoot
  [FT_Message]::execution((Split-Path -Leaf $PSCommandPath))
}

$script:order = @{
  watch_folder = $rsv_folder;
  watch_file   = $rsv_file;
  orders       = $gZEN_orders;
  action_block = $gZEN_Block;
}


[FT_Ezy_Watcher]::Start(@($order))

