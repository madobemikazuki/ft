Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

. ..\ft_cores\FT_IO.ps1
$config = [FT_IO]::Read_JSON_Object(".\config\dl.json")

# 当初、$config.targets をJSONファイルから読み込もうとしたが、
# エスケープ文字が必要なのが面倒になりテキストファイルを配列にすることで
# 手入力を簡略化した。
$dl_list = Get-Content -Path (${HOME} + $config.targets)
#$dl_list

$dist_folder = (${HOME} + $config.destination_folder)

foreach ($_p in $dl_list) {
  Write-Host "<Copied>"
  # Windows のファイルシステムでは、フォルダ名に半角 & を使用している場合、
  # 半角 $ でエスケープすることができる。すなわち、スクリプトからアクセスできる。 
  $private:from_p = $_p.replace("&", "$")

  Write-Host "　From: " $from_p
  $filename = Split-Path $from_p -Leaf
  $private:dist_path = ($dist_folder + $filename)
  
  Write-Host "　　To: " $dist_path
  Copy-Item -Path $from_p -Destination $dist_path -Force
}
Write-Host "出力完了: 💩"

