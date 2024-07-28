Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [PSCustomObject[]][ref]$_obj_list
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

function private:fn_Contains_Empty {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject]$_obj
  )
  $boolean_list = foreach ($_value in $_obj.psobject.Properties.Value) {
    [String]::IsNullOrEmpty($_value)
  }
  return ($boolean_list -contains $True)
}

  
  $incomplete_list = $_obj_list | Where-Object { fn_Contains_Empty $_ }
  return $incomplete_list 
