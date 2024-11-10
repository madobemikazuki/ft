Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [ValidatePattern("applicants|wbc|gZEN|coms|ed|wid|coh|rsv|ft_init|adding|ExcelToCSV|registered|FnD")]$_command_name
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
$out_dic = "${HOME}\Downloads"
$config_path = '.\config\deploy_files.json'
$files = @(Get-Content -Path $config_path | ConvertFrom-Json)
$empty_files_folder = ($out_dic + '\empty_files\' + $_command_name)


foreach ($file in $files.$_command_name) {
  $output_path = ($empty_files_folder + $file)
  $extension = [System.IO.Path]::GetExtension($file)
  if ($extension -eq '.txt') {
    New-Item -Path $output_path -ItemType File -Force
    Out-File -FilePath $output_path -Encoding Default -InputObject $empty -Force
    #ここで次の$fileループへ移動する
    continue
  }

  if ($extension -eq '.bat') {
    # UTF8 の bom なしエンコード
    New-Item -Path $output_path -ItemType File -Force
    # [System.IO.File]::WriteAllLines() を使用すると UTF8 BOMなしで出力できる。
    [System.IO.File]::WriteAllLines($output_path, $empty)
    continue
  }
  # utf8かつBOM付きのPowerShellスクリプトファイルを生成する。
  #Write-Host ($empty_files_dic + $file)  
  New-Item -Path $output_path -ItemType File -Force
  Out-File -FilePath $output_path -Encoding utf8 -InputObject $empty -Force
}

exit 0

