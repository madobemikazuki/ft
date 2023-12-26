Param(
  [Parameter(Mandatory = $True)]
  [String]$_corporate_name
)


Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

switch ($_corporate_name){
  {$_.Contains('株式会社')} {return $_.Replace('株式会社', '（株）')}
  {$_.Contains('有限会社')} {return $_.Replace('有限会社', '（有）')}
}

exit 0
