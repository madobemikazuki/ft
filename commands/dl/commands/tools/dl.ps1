Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

. ..\ft_cores\FT_IO.ps1
$config = [FT_IO]::Read_JSON_Object(".\config\dl.json")

# 当初、$config.targets をJSONファイルから読み込もうとしたが、
# エスケープ文字が必要なのが面倒になりテキストファイルを配列にすることで
# 手入力を簡略化した。
$dl_list = [FT_IO]::Read_JSON_Array((${HOME} + $config.targets))
#$dl_list
$dist_folder = (${HOME} + $config.destination_folder)
if (!(Test-Path  $dist_folder)){
  New-Item $dist_folder -ItemType Directory
}

foreach ($_p in $dl_list) {

  # Windows のファイルシステムでは、フォルダ名に半角 & を使用している場合、
  # 半角 $ でエスケープすることができる。すなわち、スクリプトからアクセスできる。 
  $private:target_folder = ($_p.folder).replace("&", "$")
  $private:item = Get-ChildItem -Path $target_folder -File -Filter $_p.file_name
  $private:from_path = $item.fullname
  Write-Host "　From: " $from_path

  $private:dist_path = ($dist_folder + $_p.file_name)  
  Write-Host "　　To: " $dist_path
  Copy-Item -Path $from_path -Destination $dist_path -Force
}
Write-Host "出力完了: 💩"

