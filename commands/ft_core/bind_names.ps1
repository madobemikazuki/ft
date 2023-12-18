  Param(
      [Parameter(Mandatory = $True)]
      [PSCustomObject[]]$name_list
  )

  Set-StrictMode -Version 3.0
  $ErrorActionPreference = "Stop"

  $Z_BLANC = '　'
  $UNDER_SCORE = '_' 

  $names = foreach ($name in $name_list){
      $name.replace($Z_BLANC, "")
  }
  return $names -join $UNDER_SCORE

  exit 0