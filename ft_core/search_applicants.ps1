Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [PSCustomObject[]]$_applicants,
  [Parameter(Mandatory = $True, Position = 1)]
  [ValidatePattern("^\d{2}\b\-\b\d{6}$")]
  [String[]]$_regist_nums,
  [Parameter(Mandatory = $False, Position = 2)]
  [String]$_target = "中登番号"
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

try {
  #PSCustomObjectのリスト
  #$target = '中登番号'
  #return $applicants | Select-Object | Where-Object { $_.$target -eq $regist_num }  

  $fit_peoples = foreach ($regist_num in $_regist_nums){
    $_applicants | Select-Object | Where-Object{$_.$_target -eq $regist_num}
  }
  return $fit_peoples
}
catch {
  <#Do this if a terminating exception happens#>
  Write-Host "貴様の入力した中登番号は存在しない。"
}



function show {
  $applicants | Format-Table
}
exit 0