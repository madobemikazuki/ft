Param(
  [Parameter(Mandatory = $true)]
  [ValidatePattern("^\d{7}")][String]$arg
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

return Write-Output("{0:000-0000}" -f [Int]$arg)