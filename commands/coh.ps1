# Cut off Head. 最初の一行目を削除する。
Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

try {
  $file_name = "登録状況リスト.csv"
  $target = "${HOME}\Downloads\from_T\temp\$file_name"
  $file = Get-Content -Path $target
  #$file = Get-Content -Path .\$file_name
  Set-Content -Path $target $file[1..($file.Length - 1)]
}
catch [exception] {
  Write-Output "😢😢😢エラーをよく読んでね。"
  $error[0].ToString()
  Write-Output $_
}