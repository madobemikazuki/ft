Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"


. .\ft_cores\FT_IO.ps1
$names = [FT_IO]::Read_JSON(".\config\FnD.json")

# 対象のフォルダ直下のファイルすべてを対象にするために
# $Folder(フォルダパス)にアスタリスク * を追加している。
# Get-ChildItem で -Include オプションを利用するときにワイルドカード指定が必要になる。
$private:Folder = "${HOME}\Downloads\*"
$private:head = "*"
$private:end = "*.*"

$targets = foreach ($_ in $names) {
  $head + $_ + $end
}
#$targets | Format-List

if ([FT_IO]::Exists_Path($Folder, $targets)) {
  $target_paths = (Get-ChildItem -Path $folder -File -Include $targets).FullName
  $target_paths | Format-List
  Remove-Item -Path $target_paths
  Write-Host "🌸🌸🌸 以上のファイルを 削除 しました。"
  exit 0
}

Write-Host "🌳🌳🌳 削除すべきファイルはまだ存在しません。"
exit 0

