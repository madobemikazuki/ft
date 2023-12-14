Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [String]$Folder,
  [Parameter(Mandatory = $True, Position = 1)]
  [String]$TargetName
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

try {
  $file_path_list = (Get-childItem -Path $Folder -File -Include $TargetName).fullname
  return $file_path_list
}
catch {
  Write-Host "エラー発生 :: $($_.Exception.Message)"
}

