Param(
  [Parameter(Mandatory = $True)]
  [String]$_path
)
$json  = Get-Content -Path $_path | ConvertFrom-Json
return $json
