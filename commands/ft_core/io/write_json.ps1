Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [String]$_path,
  [Parameter(Mandatory = $True, Position = 1)]
  [PSCustomObject]$_source
)

<#
  複雑なオブジェクト形式のJSONに使う。
  config.json な感じ。

#>

$utf8_with_BOM = New-Object System.Text.UTF8Encoding $True
[System.IO.File]::WriteAllLines($_path, (ConvertTo-Json $_source), $utf8_with_BOM)