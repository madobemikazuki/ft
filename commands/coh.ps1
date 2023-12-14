<#
使用するフォルダに coh.ps1 ファイルを移動し、
実行すること
#>

# Cut off Head. 最初の一行目を削除する。
Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"


$file_name = "登録者管理リスト.csv"
#$file = Get-Content -Path ${HOME}\Downloads\from_T\temp\$file_name
$file = Get-Content -Path .\$file_name
Set-Content -Path .\$file_name $file[1..($file.Length - 1)]