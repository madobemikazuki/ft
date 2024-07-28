function Contains-Empty {
  Param(
    [CmdletBinding()]
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject]$_obj
  )

  Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"
  $boolean_list = foreach ($_value in $_obj.psobject.Properties.Value) {
    [String]::IsNullOrEmpty($_value)
  }
  return ($True -in $boolean_list)
}
