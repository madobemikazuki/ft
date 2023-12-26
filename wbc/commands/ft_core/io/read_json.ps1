Param(
  [Parameter(Mandatory = $True)]
  [ValidatePattern('\.json$')]$_path
)
$json  = Get-Content -Path $_path | ConvertFrom-Json
return $json
