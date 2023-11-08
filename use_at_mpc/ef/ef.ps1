Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [ValidatePattern("wbc")]$_command_name
)


Set-StrictMode -Version 3.0

<#
.SYNOPSIS
  ftで用いているコマンド群の フォルダ と 空の.ps1ファイル を生成する。
.DESCRIPTION
  ローカルにスクリプト群をコピペする時に少しでも楽できたらなぁと思った。
.EXAMPLE
 PS> . .\ef.ps1 wbc
.INPUTS
  wbc
.OUTPUTS
  フォルダと.ps1ファイルを出力する。
.NOTES
  生成後の .ps1 ファイルは全て空になっている。
  UTF-8 with BOM になる。
.COMPONENT
  このコマンドレットが属するコンポーネント
.ROLE
  このコマンドレットが属する役割
#>




$empty = ""
$out_dic ="${HOME}\Downloads"
$config_path = '.\config\script_files.json'
$files = @(Get-Content -Path $config_path | ConvertFrom-Json)
$empty_files_dic = ($out_dic + '\empty_files\' + $_command_name)


foreach ($file in $files.$_command_name) {
  # utf8かつBOM付きのPowerShellスクリプトファイルを生成する。
  #Write-Host ($empty_files_dic + $file)
  $destination = ($empty_files_dic + $file)
  New-Item -Path $destination -ItemType File -Force
  Out-File -FilePath $destination -Encoding utf8 -InputObject $empty -Force
}

exit 0