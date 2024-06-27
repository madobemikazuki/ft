Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [PSCustomObject[]]$_applicants,
  [Parameter(Mandatory = $True, Position = 1)]
  [String[]]$_printing_field,
  [Parameter(Mandatory = $True, position = 2)]
  [PSCustomObject[]]$_address_table
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"


function private:fn_mapping {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject]$_applicant,
    [Parameter(Mandatory = $True, Position = 1)]
    [String[]]$_field,
    [Parameter(Mandatory = $True, Position = 2)]
    [PSCustomObject]$_address
  )
  foreach ($_name in $_field) {
    [PSCustomObject]@{
      Name     = $_name
      Value    = $_applicant.$_name
      point_x = $_address.$_name[0]
      point_y = $_address.$_name[1]
    }
  }
}


$private:applicant_list = $_applicants
#申請者情報をExcel 転記用のフォーマットに変換する
$formated_obj_list = foreach ($_applicant in $applicant_list) {
  $index = $_applicants.indexOf($_applicant)
  $address = fn_mapping $_applicant $_printing_field $_address_table[$index]
  $address
}
return $formated_obj_list
