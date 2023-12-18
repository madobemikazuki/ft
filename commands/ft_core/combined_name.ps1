Param(
  [Parameter(Mandatory = $true, Position=0)][String]$first_name,
  [Parameter(Mandatory = $true, Position=1)][String]$last_name,
  [String]$delimiter = '　'#デフォルト引数 呼び出し側で -delimiter を指定すること
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

$sb = New-Object System.Text.StringBuilder
#副作用処理  StringBuilderならちょっと速いらしい。要素数が少ないから意味ないかも。
@($first_name, $delimiter , $last_name) | ForEach-Object { [void] $sb.Append($_) }
return $sb.ToString()
