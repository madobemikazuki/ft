Param(
  [Parameter(Mandatory = $True, Position = 0)][PSCustomObject[]][ref]$_obj_list
)

<#
  PSCustomObject の ヘッダーのみを返す
#>

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

$_obj_list[0] | Get-Member -Membertype NoteProperty | Select-Object -ExpandProperty Name
