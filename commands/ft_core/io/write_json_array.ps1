Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [String]$_path,
  [Parameter(Mandatory = $True, Position = 1)]
  [PSCustomObject[]]$_source
)

<#
  こんな形で保存するのに使う
  JavaScript でも使えればいい。
  [
    "文字列",
    "文字列"
  ]
#>

$utf8_with_BOM = New-Object System.Text.UTF8Encoding $True
[System.IO.File]::WriteAllLines($_path, (ConvertTo-Json $_source), $utf8_with_BOM)