<#
  申請会社や主管グループ等を収録するJSONを出力するスクリプト
#>
function private:fn_Extract_Set {
  Param(
    [Parameter(Mandatory = $True)]
    [PSCustomObject[]][ref]$psc_list,
    [Parameter(Mandatory = $True)]
    [String[]][ref]$keys
  )
  $private:coms = . .\commands\utils\ol.ps1 map $psc_list $keys
  $private:coms_set = . .\commands\utils\ol.ps1 unique $coms $keys[0]
  return $coms_set
}

function initial_ZERO {
  Param(
    [Parameter(Mandatory = $True)]
    [String]$_str
  )
  return $_str -replace "/^T/","0"
}

function private:fn_Create_KV{
  Param(
    [Parameter(Mandatory = $True)]
    [PSCustomObject[]][ref]$_psobj_list,
    [Parameter(Mandatory = $True)]
    [String][ref]$_key,
    [Parameter(Mandatory = $True)]
    [String][ref]$_value
  )
  $private:kv = [PSCustomObject]@{}
  foreach ($_ in $_psobj_list){
    $kv | Add-Member -MemberType NoteProperty -Name $_.$_key -Value $_.$_value
  }
  return $kv
}

# 登録にも解除にも必要な情報
[PSCustomObject[]]$private:source = . .\commands\ft_core\io\read_registed_people_fromT.ps1
[String[]]$app_coms_field = @('電力申請会社番号', '電力申請会社名称')
[PSCustomObject[]]$temp_coms_set = fn_Extract_Set ([ref]$source) ([ref]$app_coms_field)

[PSCustomObject[]]$app_coms_set = foreach ($key in $temp_coms_set.psobject.properties.name){
  $_.psobject.properties.name 
}

#$app_coms_set | Format-Table
$app_coms_kv = fn_Create_KV ([ref]$app_coms_set) ([ref]$app_coms_field[0]) ([ref]$app_coms_field[1])
#$app_coms_kv


# 登録時に必要になる。
$emp_coms_feild = @('グループ', '雇用番号', '雇用名称')
#$emp_coms_set = $source | Select-Object -Property $emp_coms_feild
$emp_coms = fn_Extract_Set ([ref]$source) ([ref]$emp_coms_feild)
#$emp_coms | ft
#登録に必要になる。かな？
[PSCustomObject[]]$group_set = fn_Extract_Set ([ref]$emp_coms) ([ref]$emp_coms_feild[0])


$coms_object = [PSCustomObject]@{
  AppComs = $app_coms_kv
  EmpComs = $emp_coms
  TGroup = foreach($_ in $group_set){$_.psobject.properties.value }
}

$path = "${HOME}\Downloads\TEMP\coms.json"
. .\commands\ft_core\io\write_json.ps1 $path $coms_object

. .\commands\utils\notify_balloon.ps1 'coms' "出力完了🌲🌲🌲 : $path"
exit 0