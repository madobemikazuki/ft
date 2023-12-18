Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [ValidatePattern("wbc|gZEN")]$_command_name
)

Set-StrictMode -Version 3.0
$source_folder = "${HOME}\apps\ft\commands"
$config_path = "$source_folder\config\deploy_files.json"
$script_files = @(Get-Content -Path $config_path | ConvertFrom-Json)

$_command_name
# $script_files.$_command_name

$destination_folder = "${HOME}\Downloads\output\toGitHub\$_command_name"

#$destination_folder

foreach ($file in $script_files.$_command_name) {
  if ([System.IO.Path]::GetExtension($file) -eq '.txt'){
    $text = Get-Content -Path "$source_folder$file" -raw -Encoding Default
    New-Item -Path "$destination_folder$file" -ItemType File -Force
    # CSVのヘッダーに使用するテキストファイルの末尾に改行文字があるとエラーになる。
    # 格納先にnullが渡されてしまうので、$textを TrimEnd()処理する必要がある。
    $trimed_text = $text.TrimEnd()
    Out-File -FilePath "$destination_folder$file" -Encoding Default -InputObject $trimed_text -Force
    #ここで次の次のループへ移動する
    continue
  }
  $text = Get-Content -Path "$source_folder$file" -raw -Encoding utf8
  New-Item -Path "$destination_folder$file" -ItemType File -Force
  Out-File -FilePath "$destination_folder$file" -Encoding utf8 -InputObject $text -Force
}



#$extension = ".json"
##$config_path = "${HOME}\Downloads\config\$_command_name$extension"
#New-Item -Path "$destination_folder\config" -ItemType Directory -Force
#Copy-Item $config_path "$destination_folder\config\$_command_name$extension" -Force