Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

. ..\ft_cores\FT_IO.ps1

$config = [FT_IO]::Read_JSON_Object(".\config\dl.json")
$dl_list = [FT_IO]::Read_JSON_Array((${HOME} + $config.plist))
#$dl_list

$dist_folder = (${HOME}+ $config.destination_folder)

foreach ($_p in $dl_list){
  Write-Host "Copied..."
  # Windows のファイルシステムでは、フォルダ名に半角 & を使用している場合、
  # 半角 $ でエスケープすることができる。すなわち、スクリプトからアクセスできる。 
  $from_p = $_p.replace( "&","$")
  Write-Host "　From: " $from_p

  $filename = Split-Path $from_p -Leaf
  $dist_path = ($dist_folder + $filename)
  Write-Host "　　To: " $dist_path

  Copy-Item -Path $from_p -Destination "$dist_path" -Force
}
Write-Host "出力完了: 💩"

