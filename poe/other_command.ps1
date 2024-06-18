Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

function fn_export_file_names {
  Param(
    [Parameter(Mandatory = $true, Position = 0)]
    [String[]]$_names,
    [Parameter(Mandatory = $True, Position = 1)]
    [String]$_extension
  )
  $private:joined_name = $_names -join "_"
  $private:trimed_name = $joined_name -replace "　", ""
  return  ($trimed_name + $_extension)
}

$config_path = ".\config\poe.json"
. .\core\read.ps1
$config = [PSCustomObject](fn_read $config_path)
#$config.poe_config | Format-List

$reserved_info_path = (${HOME} + $config.reserv_info_path)
$reserved_info_list = [PSCustomObject[]] (fn_read $reserved_info_path)
#$reserved_info_list | Format-Table
$required_info_list = $reserved_info_list | Select-Object -Property $config.extraction_list
#$required_info_list | Format-Table

#テストデータを参照 コマンドラインからの入力を想定
$regist_nums = @("30-394422", "14-701576", "15-977244", "77-271139", "31-107099")
$target_list = foreach($_num in $regist_nums){
  $required_info_list | Where-Object {$_."中央登録番号" -eq $_num }
}
$target_list | Format-Table

$names = foreach ($_ in $required_info_list) { $_."漢字氏名" }
$file_names = ($config.export_file_name_head + (fn_export_file_names $names ".xls"))
#$file_names

$export_folder = (${HOME} + $config.export_folder)
$export_path = ($export_folder + $file_names)
$export_path

. .\poe.ps1 $config.poe_config $target_list $export_path

exit 0
