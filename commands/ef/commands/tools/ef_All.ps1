Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

. ..\ft_cores\FT_IO.ps1

$config = [FT_IO]::Read_JSON_Object(".\config\\deploy_files.json")
$keys = $config.psobject.properties.Name

foreach($_key in $keys){
  . .\ef.ps1 $_key
}
Write-Host "空ファイル群を作成しました。"

