function Excluded-List {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject[]]$_obj_list,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject[]]$_exclusion_list,
    [Parameter(Mandatory = $True, Position = 2)]
    [String][ref]$_key
  )
  Set-StrictMode -Version 3.0
  $ErrorActionPreference = "Stop"

  $exclusion_value_list = foreach ($_ in $_exclusion_list) { $_.$_key }
  # obj_list に $_value_list の要素と一致しないリストを返す
  $result = $_obj_list | Where-Object { $exclusion_value_list -notcontains $_.$_key }
  return $result
}

