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
  $_obj_list | Where-Object { $_.$_prop -eq $_value }
}

function script:fn_Insert_Reserved {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject[]]$_applicant_list,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject[]]$_reserved_info_list
  )
  $key = '中央登録番号'
  $addition_keys = @("登録予約日", "登録予約時間", "管理会社")

  $inserted_applicant_list = foreach ($_applicant in $_applicant_list) {
    # 元データの中央登録番号が空だとスクリプトがここで停止するので対策した。
    # $_applicant.$keyが空のときは次の要素へスキップする。
    if ([String]::IsNullOrEmpty($_applicant.$key)) { continue }
    $reserved_info = fn_Array_Filter $_reserved_info_list $key $_applicant.$key
    #Write-Host $reserved_info
    #Write-Host $reserved_info.GetType()
    $inserted_applicant = fn_Append_KV $_applicant $addition_keys $reserved_info
    $inserted_applicant
  }
  return $inserted_applicant_list
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
    #$_obj | Add-Member -NotePropertyName $_key -NotePropertyValue $_addition.$_key -Force
    Add-Member -InputObject $_obj -NotePropertyName $_key -NotePropertyValue $_addition.$key -Force
    #$_obj | Add-Member -NotePropertyMembers @{$_key = $_addition.$_key } -Force
  }
  return $_obj
}

function script:fn_Shorten_Com_Type_Name {
  Param(
    [Parameter(Mandatory = $True)]
    [String]$_corporate_name
  )
  switch ($_corporate_name) {
    { $_ -match "株式会社|\(株\)" } { return $_ -replace "株式会社|\(株\)", "（株）" }
    { $_ -match "有限会社|\(有\)" } { return $_ -replace "有限会社|\(有\)", "（有）" }
    default { return $_ }
  }
}


function script:fn_Application_Company_Names {
  Param(
    [Parameter(mandatory = $True, Position = 0)]
    [PSCustomObject]$_applicant
  )
  $application_com = fn_Shorten_Com_Type_Name $_applicant.'所属企業名' 
  $manegement_com = fn_Shorten_Com_Type_Name $_applicant.'管理会社'
  $employment_com = if ([String]::IsNullOrEmpty($_applicant.'雇用企業名称（漢字）')) {
    fn_Shorten_Com_Type_Name $_applicant.'所属企業名' 
  }
  else { fn_Shorten_Com_Type_Name $_applicant.'雇用企業名称（漢字）' }

  # TODO: 派遣が含まれている場合、$_applicant.'管理会社'のみでよいかもしれない。要検証。
  if ($manegement_com.Contains("派遣")) { return $manegement_com }
  if ($application_com -eq $employment_com) { return $application_com }
  return ($application_com + "／" + $employment_com)
}


<#
  ここから下は実行内容
#>
Write-Host "登録 予約済み情報の出力"

$script:config = fn_Read ".\config\bind_r.json"

$gZEN_Path = (${HOME} + $config.gZEN_Path)
[PSCustomObject[]]$script:gZEN_obj_list = Get-Content -Path $gZEN_Path -Encoding UTF8 | ConvertFrom-Json
#$gZEN_obj_list

$reserved_Path = (${HOME} + $config.reserved_Path)
[PSCustomObject[]]$script:reserved_obj_list = fn_Read $reserved_Path
#$reserved_obj_list

[PSCustomObject[]]$binded_list = fn_Insert_Reserved $gZEN_obj_list $reserved_obj_list
#$binded_list


# TODO: 副作用 どうにかしたい
# 元のプロパティに再代入するのではなく、新たなプロパティに必要な値を代入する。
foreach ($_ in $binded_list) {
  $app_com_names = fn_Application_Company_Names $_
  Add-Member -InputObject $_ -NotePropertyName $config.selection_key -NotePropertyValue $app_com_names -Force
}

$selection = $binded_list | Select-Object -Property $config.r_selection
$selection |Format-Table


$utf8_BOM = New-Object System.Text.UTF8Encoding $True
$json_path = (${HOME} + $config.export_Path)
fn_Write_JSON $json_path $selection $utf8_BOM

