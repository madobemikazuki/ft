Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [PSCustomObject[]]$_obj_list,
  [Parameter(Mandatory = $True, position = 1)]
  [PSCustomObject]$_address_table
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"


function fn_mapping {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject]$_applicant,
    [Parameter(Mandatory = $True, Position = 1)]
    [String[]]$_field,
    [Parameter(Mandatory = $True, Position = 2)]
    [PSCustomObject]$_address,
    [Parameter(Mandatory = $True, Position = 3)]
    [byte]$_index
  )
  foreach ($_name in $_field) {
    [PSCustomObject]@{
      Name     = $_name
      Value    = $_applicant.$_name
      point_x = ($_address.$_name[0] + $_index)
      point_y = $_address.$_name[1]
    }
  }
}

$field = $_address_table.PSObject.Properties.Name

#申請者情報をExcel 転記用のフォーマットに変換する
$formated_obj_list = foreach ($_obj in $_obj_list) {
  $index = $_obj_list.indexOf($_obj)
  $address = fn_mapping $_obj $field $_address_table $index

  $address
}
return $formated_obj_list

