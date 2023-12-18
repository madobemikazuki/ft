Param(
  [Parameter(Mandatory = $true, Position = 0)]
  [String[]] $_header,
  [Parameter(Mandatory = $true, Position = 1)]
  [PSCustomObject]$_position,
  [Parameter(Mandatory = $true, Position = 2)]
  [PSCustomObject]$_applicant
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

$COLON = ':'

[Array]$address = foreach ($_ in $_header) {
  $p = [UInt16[]] $_position.$_.split($COLON)
  [PSCustomObject] @{
    name    = $_
    point_x = $p[0]
    point_y = $p[1]
    value   = $_applicant.$_
  }
}
return $address
