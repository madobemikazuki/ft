Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [String]$_path,
  [Parameter(Mandatory = $True, Position = 1)]
  [PSCustomObject[]]$_Object_List,
  [Parameter(Mandatory = $True, Position = 2)]
  [System.Text.Encoding]$_encoding
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"


if (Test-Path $_path) {
  New-Item -Path $_path -ItemType File -Force
}
[System.IO.File]::WriteAllLines($_path, (ConvertTo-Json $_Object_List), $_encoding)
