Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [PSCustomObject]$_poe_config,
  [Parameter(Mandatory = $True, Position = 1)]
  [PSCustomObject[]]$_info_obj_list,
  [Parameter(Mandatory = $True, Position = 2)]
  [String]$_export_path
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

if($_poe_config.printing.style -eq "chunk"){
  Write-Host "チャンク転記処理するよ"
  exit 0
}

if($_poe_config.printing.style -eq "single"){
  Write-Host "シングル転記処理するよ"
  exit 0
}

exit 0
