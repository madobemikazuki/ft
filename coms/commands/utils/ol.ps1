Param(
  [Parameter(Mandatory = $True, Position = 0)][ValidatePattern('map|filter|search|sort|unique|inquiry')]$Task,
  [Parameter(Mandatory = $True, Position = 1)][PSCustomObject[]][ref]$_obj_list,
  [Parameter(Mandatory = $True, Position = 2)][String[]][ref]$_keys,
  [Parameter(Mandatory = $False, Position = 3)][String[]][ref]$_targets
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"


# Object List のコレクション操作

function fn_map {
  Param(
    [Parameter(Mandatory = $True, Position = 0)][PSCustomObject[]][ref]$_obj_list,
    [Parameter(Mandatory = $True, Position = 1)][String[]][ref]$_keys
  )
  $_obj_list | Select-Object -Property $_keys
}


function fn_sort {
  Param(
    [Parameter(Mandatory = $True, Position = 0)][PSCustomObject[]][ref]$_obj_list,
    [Parameter(Mandatory = $True, Position = 1)][String[]][ref]$_keys
  )
  $key = $_keys[0]
  $_obj_list | Sort-Object -Property $key
}


function fn_search {
  Param(
    [Parameter(Mandatory = $True, Position = 0)][PSCustomObject[]][ref]$_obj_list,
    [Parameter(Mandatory = $True, Position = 1)][String[]][ref]$_keys,
    [Parameter(Mandatory = $True, Position = 2)][String[]][ref]$_targets
  )
  $key = $_keys[0]
  foreach ($target in $_targets) {
    $_obj_list.Where({ $_.$key -eq $target })
  }
}


function fn_unique {
  Param(
    [Parameter(Mandatory = $True, Position = 0)][PSCustomObject[]][ref]$_obj_list,
    [Parameter(Mandatory = $True, Position = 1)][String[]][ref]$_keys
  )
  fn_sort $_obj_list $_keys | Get-Unique -AsString
}


function fn_filter {
  Param(
    [Parameter(Mandatory = $True, Position = 0)][PSCustomObject[]][ref]$_obj_list,
    [Parameter(Mandatory = $True, Position = 1)][String[]][ref]$_keys,
    [Parameter(Mandatory = $True, Position = 2)][String[]][ref]$_values
  )

  <#
  複数のパラメータで検索し抽出する
  $_keys と $_values が 同じ長さであること
  fn_search を併用すると簡潔に書けるかも
  #>
  $key = $_keys[0]
  foreach ($target in $_targets) {
    $_obj_list.Where({ $_.$key -eq $target })
  }

}

function fn_inquiry{
  Param(
    [Parameter(Mandatory = $True, Position = 0)][PSCustomObject[]][ref]$_obj_list,
    [Parameter(Mandatory = $True, Position = 1)][String][ref]$_key,
    [Parameter(Mandatory = $True, Position = 2)][String][ref]$_value
  )
  #リストの中から一つのkeyを指定し、valueに該当するものを返す
  $_obj_list.Where({$_.$_key -eq $_value})
}



switch ($Task) {
  map { fn_map $_obj_list $_keys }
  sort { fn_sort $_obj_list $_keys }
  filter { fn_filter  $_obj_list $_keys $_targets }
  search { fn_search  $_obj_list $_keys $_targets }
  unique { fn_unique $_obj_list $_keys }
  inquiry { fn_inquiry $_obj_list $_keys[0] $_targets[0]}
}
