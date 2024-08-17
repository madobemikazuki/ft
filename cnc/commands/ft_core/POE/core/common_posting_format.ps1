Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [PSCustomObject]$_common_obj,
  [Parameter(Mandatory = $True, position = 1)]
  [PSCustomObject]$_address_table
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"


$field = $_common_obj.PSObject.Properties.Name

#共通情報を転記用のフォーマットに変換する
$new_obj_list = foreach ($_ in $field) {
  [PSCustomObject]@{
    Name    = $_
    Value   = $_common_obj.$_
    point_x = $_address_table.$_[0]
    point_y = $_address_table.$_[1]
  }
}

return $new_obj_list

