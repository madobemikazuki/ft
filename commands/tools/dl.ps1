Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

. ..\ft_cores\FT_IO.ps1
. ..\ft_cores\FT_Message.ps1

$private:config = [FT_IO]::Read_JSON_Object(".\config\dl.json")
[FT_Message]::execution($config.command_name)
# 当初、$config.targets をJSONファイルから読み込もうとしたが、
# エスケープ文字が必要なのが面倒になりテキストファイルを配列にすることで
# 手入力を簡略化した。
$private:dl_list = [FT_IO]::Read_JSON_Array((${HOME} + $config.targets))
#$dl_list
$private:dist_folder = (${HOME} + $config.destination_folder)
if (!(Test-Path  $dist_folder)){
  New-Item $dist_folder -ItemType Directory
}

foreach ($_path in $dl_list) {
  # Windows のファイルシステムでは、フォルダ名に半角 & を使用している場合、
  # 半角 $ でエスケープすることができる。すなわち、スクリプトからアクセスできる。 
  $private:target_folder = ($_path.folder) -replace "&|＆","$"
  $private:item = Get-ChildItem -Path $target_folder -File -Filter $_path.file_name
  $private:from_path = $item.fullname
  Write-Host "　From: " $from_path

  $private:dist_path = ($dist_folder + $_path.file_name)  
  Write-Host "　　To: " $dist_path
  Copy-Item -Path $from_path -Destination $dist_path -Force
}
Write-Host "出力完了: 💩"
