set-Variable COLON ':' -option Constant

function map_address {
  Param(
    [Parameter(Mandatory = $true)]
    [String[]] $header,
    [PSCustomObject]$position,
    [PSCustomObject]$applicant
  )
  [Array]$a = foreach ($_ in $header) {
    $p = [UInt16[]] $position.$_.split($COLON)
    [PSCustomObject] @{
      name    = $_
      point_x = $p[0]
      point_y = $p[1]
      value   = $applicant.$_
    }
  }
  return $a
}


# Excel シートへ転記するアドレス情報と転記必須情報を統合する。
function combine_objects {
  Param(
    [Parameter(Mandatory = $true)]
    [PSCustomObject]$address_table,
    [PSCustomObject]$mandatory_table
  )
  $items = foreach ($_ in $address_table.psobject.properties.name) {
    $point = [UInt16[]] $address_table.$_.split($COLON)
    [PSCustomObject] @{
      field   = $_
      point_x = $point[0]
      point_y = $point[1]
      value   = $mandatory_table.$_
    }
  }
  return $items
}