Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

. ..\ft_cores\FT_IO.ps1
. ..\ft_cores\FT_Path.ps1
. ..\ft_cores\FT_Message.ps1

$private:config = [FT_IO]::Read_JSON_Object(".\config\dl.json")
[FT_Message]::execution($config.command_name)
# 当初、$config.targets をJSONファイルから読み込もうとしたが、
# エスケープ文字が必要なのが面倒になりテキストファイルを配列にすることで
# 手入力を簡略化した。
$private:dl_list = [FT_IO]::Read_JSON_Array((${HOME} + $config.targets))
$dl_list | Format-Table
$private:fixed_dist_folder = [FT_Path]::Fixed_Path($config.destination_folder)
if (!(Test-Path  $fixed_dist_folder)) {
  New-Item $fixed_dist_folder -ItemType Directory
}

foreach ($_path_obj in $dl_list) {
  # Windows のファイルシステムでは、フォルダ名に半角 & を使用している場合、
  # 半角 $ でエスケープすることができる。すなわち、スクリプトからアクセスできる。 


  # どちらがいいのやら。\TEMP\DL_targets.json を忘れずに。


  #$private:target_folder = ($_path.folder) -replace "＆","$"

  #$private:target_folder = [FT_Path]::Fixed_Path($_path.folder)
  #$private:item = Get-ChildItem -Path $target_folder -File -Filter $_path.file_name
  <#
  $private:item = Get-ChildItem -Path $_path.folder -File -Filter $_path.file_name
  $private:from_path = $item.fullname
  Write-Host "　From: " $from_path

  $private:dist_path = ($fixed_dist_folder
 + $_path.file_name)  
  Write-Host "　　To: " $dist_path
  Copy-Item -Path $from_path -Destination $dist_path -Force
#>

  $target_path = $_path_obj.folder + "\" + $_path_obj.file_name  
  #$target_path = @($_path_obj.psobject.properties.value).Join("\")
  [FT_Path]::CopyToDL($target_path, $fixed_dist_folder)
}
Remove-Variable dl_list, fixed_dist_folder


