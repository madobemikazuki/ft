Param(
  [Parameter(Mandatory = $True, Position = 0)][String]$_export_path,
  [Parameter(Mandatory = $True, Position = 1)][PSCustomObject[]]$_obj_list
)

<#
  PSCustomObject の ヘッダーのみを返す
#>

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

$header = . .\utils\header.ps1 ([ref]$_obj_list)


.\ft_core\io\write_json_array.ps1 $_export_path $header
exit 0