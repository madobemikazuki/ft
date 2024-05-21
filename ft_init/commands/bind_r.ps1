Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"


<#
機能 : gZEN_exported.gzen と 登録_予約日リスト_UTF8-bom.json の情報をバインドする。
目的 : 登録申請書に添付するWBC受検用紙の出力、
       従事者登録後のcd教育受検用紙出力、
       上記二点に必要な情報をJSON形式で出力する。
#>

function private:fn_Read {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [ValidatePattern("\.csv$|\.json$")]$_path
  )
  switch -Regex ($_path) {
    "\.csv$" {
      return Import-Csv -Path $_path -Encoding Default
    }
    "\.json$" {
      return Get-Content -Path $_path -Encoding UTF8 | ConvertFrom-Json
    }
    Default {
      Write-Host "拡張子が該当しないので終了。"
      exit 0
    }
  }
}

function script:fn_Write_JSON {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$_path,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject[]]$_Object_List,
    [Parameter(Mandatory = $True, Position = 2)]
    [System.Text.Encoding]$_encoding
  )
  if (Test-Path $_path) {
    New-Item -Path $_path -ItemType File -Force
  }
  [System.IO.File]::WriteAllLines($_path, (ConvertTo-Json $_Object_List), $_encoding)
}

function script:fn_Array_Filter {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject[]]$_obj_list,
    [Parameter(Mandatory = $True, Position = 1)]
    [String]$_prop,
    [Parameter(Mandatory = $True, Position = 2)]
    [String]$_value
  )
  #$_obj_list | Where-Object { $_.$_prop -eq $_value }
  $target_obj = foreach ($_ in $_obj_list) {
    if ($_.$_prop -eq $_value) {
      return $_
    }
  }
  return $target_obj
}

function script:fn_Insert_Reserved {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject[]][ref]$_applicant_list,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject[]][ref]$_reserved_info_list
  )
  $reg_num = '中央登録番号'
  $new_key_list = @("登録予約日", "登録予約時間", "管理会社")

  $extracted = foreach ($_applicant in $_applicant_list) {
    $_reserved_applicant_list = fn_Array_Filter $_reserved_info_list $reg_num $_applicant.$reg_num
    fn_Append_KV $_applicant $new_key_list $_reserved_applicant_list
  }
  return $extracted
}

function script:fn_Append_KV {
  Param(
    [Parameter(mandatory = $True, Position = 0)]
    [PSCustomObject]$_obj,
    [Parameter(Mandatory = $True, Position = 1)]
    [String[]]$_keys,
    [Parameter(Mandatory = $True, Position = 2)]
    [PSCustomObject]$_addition
  )
  foreach ($_key in $_keys) {
    Add-Member -InputObject $_obj -NotePropertyName $_key -NotePropertyValue $_addition.$_key -Force
  }
  return $_obj
}

function script:fn_Shorten_Com_Type_Name {
  Param(
    [Parameter(Mandatory = $True)]
    [String]$_corporate_name
  )
  switch ($_corporate_name) {
    { $_.Contains('株式会社') } { return $_.Replace('株式会社', '（株）') }
    { $_.Contains('有限会社') } { return $_.Replace('有限会社', '（有）') }
    default { return $_corporate_name }
  }
}

function script:fn_Application_Company_Names {
  Param(
    [Parameter(mandatory = $True, Position = 0)]
    [PSCustomObject]$_applicant
  )
  $application_com = $_applicant.'所属企業名' 
  $manegement_com = $_applicant.'管理会社'
  $employment_com = if ([String]::IsNullOrEmpty($_applicant.'雇用企業名称（漢字）')) {
    $_applicant.'所属企業名' 
  }
  else { $_applicant.'雇用企業名称（漢字）' }

  # TODO: 派遣が含まれている場合、$_applicant.'管理会社'のみでよいかもしれない。要検証。
  if ($manegement_com.Contains("派遣")) { return fn_Shorten_Com_Type_Name $manegement_com }
  if ($application_com -eq $employment_com) { return fn_Shorten_Com_Type_Name $application_com }
  return fn_Shorten_Com_Type_Name ($application_com + "／" + $employment_com)
}


<#
  ここから下は実行内容
#>
Write-Host "登録 予約済み情報の出力"

$script:config = fn_Read ".\config\bind_r.json"

$gZEN_Path = (${HOME} + $config.gZEN_Path)
[PSCustomObject[]]$script:gZEN_obj_list = Get-Content -Path $gZEN_Path -Encoding UTF8 | ConvertFrom-Json
#$gZEN_obj_list.Length

$reserved_Path = (${HOME} + $config.reserved_Path)
[PSCustomObject[]]$script:reserved_obj_list = fn_Read $reserved_Path
#$reserved_obj_list

[PSCustomObject[]]$binded_list = fn_Insert_Reserved ([ref]$gZEN_obj_list) ([ref]$reserved_obj_list)
#$binded_list.Length
$binded_list


# TODO: 副作用 どうにかしたい
# 元のプロパティに再代入するのではなく、新たなプロパティに必要な値を代入する。
foreach ($_ in $binded_list) {
  $app_com_names = fn_Application_Company_Names $_
  Add-Member -InputObject $_ -NotePropertyName $config.selection_key -NotePropertyValue $app_com_names -Force
}

$selection = $binded_list | Select-Object -Property $config.r_selection
#$selection


$utf8_BOM = New-Object System.Text.UTF8Encoding $True
$json_path = (${HOME} + $config.export_Path)
fn_Write_JSON $json_path $selection $utf8_BOM

