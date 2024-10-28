<#
ソースとなるcsvファイルの先頭行を削除し、
別名csvファイルに書き出す。
ソースとなったcsvファイルは別フォルダへ移動する。
#>

# Cut off Head. 最初の一行目を削除する。
Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

. ..\ft_cores\FT_IO.ps1

$command_name = Split-Path -Leaf $PSCommandPath
$config_path = ".\config\FT_Utils.json"
$config = ([FT_IO]::Read_JSON_Object($config_path)).$command_name

$source_path = (${HOME} + $config.source_path)
$text = Get-Content -Path $source_path -Encoding Default
$export_csv_path = (${HOME} + $config.export_path)
# $text の3行目から $textの最終行まで取得して、作成したファイルに書き込む

$text_list = $text[2..($text.Length - 1)]
Set-Content -Path $export_csv_path $text_list -Encoding Default
Write-Host "CutOff Head...完了 : 🐶 🐶 🐶 "

$waste_folder = (${HOME} + $config.waste_folder)
$not_exisits_waste_folder = !(Test-Path $waste_folder)
# $source_path は不要なので移動する
if($not_exisits_waste_folder){
  Write-Host "\waste\csvフォルダを作る"
  New-Item -Path $waste_folder -ItemType Directory -Force
}
Move-Item -Path $source_path -Destination $waste_folder -Force


