<#
使用するフォルダに coh.ps1 ファイルを移動し、
実行すること

\Downloads\TEMP\に配置することを想定
#>

# Cut off Head. 最初の一行目を削除する。
Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"


$target_file_name = "${HOME}\Downloads\登録者管理リスト.csv"
$file = Get-Content -Path $target_file_name -Encoding Default

# TODO: 誤って上記ファイルから何度も一行目を削除しないようにするため別名ファイルに書き出す。
$output_file_name = "${HOME}\Downloads\登録者管理リスト_coh.csv"
Set-Content -Path $output_file_name $file[1..($file.Length - 1)] -Encoding Default

