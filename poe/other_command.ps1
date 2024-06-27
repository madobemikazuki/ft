Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"


$config_path = ".\config\poe.json"
. .\core\read.ps1
$config = [PSCustomObject](fn_read $config_path)
#$config.poe_config | Format-List

$reserved_info_path = (${HOME} + $config.reserv_info_path)
$reserved_info_list = [PSCustomObject[]] (fn_read $reserved_info_path)
#$reserved_info_list | Format-Table

# required == 必須  必須な情報だけ抽出する。
$required_info_list = $reserved_info_list | Select-Object -Property $config.extraction_list
#$required_info_list | Format-Table

#テストデータを参照 コマンドラインからの入力を想定
$regist_nums = @()
$target_list = foreach($_num in $regist_nums){
  $required_info_list | Where-Object {$_."中央登録番号" -eq $_num }
}
$target_list | Format-Table

. .\poe.ps1 $config.poe_config $target_list

exit 0
