Param(
  [Parameter(Mandatory = $true, Position = 0)][Object[]]$_csv_obj,
  [Parameter(Mandatory = $true, Position = 1)][String]$_path,
  [String]$_delimiter = ',',
  [String]$_encode = "UTF8"# Default ではブラウザで参照すると文字化けする。
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"
#$_path
$_csv_obj | Export-Csv -NotypeInformation -Path $_path -Delimiter $_delimiter -Encoding $_encode -Force