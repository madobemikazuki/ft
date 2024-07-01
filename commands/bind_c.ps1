Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"


<#
機能 : 登録者管理リスト_coh.csv と 解除_予約日リスト_UTF8-bom.json の情報をバインドする。
目的 : 解除申請書に添付するWBC受検用紙の出力、
       解除申請書の用紙出力、
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
  # $_obj_list | Where-Object { $_.$_obj_prop -eq $_value }
  $target_obj = foreach ($_obj in $_obj_list) {
    if ($_obj.$_prop -eq $_value) {
      $_obj
    }
  }
  return $target_obj
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
    { $_ -match "株式会社|\(株\)" } { return $_ -replace "株式会社|\(株\)", "（株）" }
    { $_ -match "有限会社|\(有\)" } { return $_ -replace "有限会社|\(有\)", "（有）" }
    default { return $_ }
  }
}
function script:fn_To_Wide {
  Param(
    [Parameter(Mandatory = $True)][String]$half_string
  )
  Add-Type -AssemblyName "Microsoft.VisualBasic"
  [Microsoft.VisualBasic.Strings]::StrConv($half_string, [Microsoft.VisualBasic.VbStrConv]::Wide)
}

function script:fn_WBC_Company_Names {
  Param(
    [Parameter(mandatory = $True, Position = 0)]
    [PSCustomObject][ref]$_applicant
  )
  # この処理は bind_r.ps1 とは若干違う。
  # 理由: 登録者管理リスト_coh.csv のフィールドと事前申請のフィールドが異なるため。
  # if で判定する前に企業名を fn_Shorten_Com_Type_Nameするのは
  # ソースとなる Companies.json に半角全角の(株)（株）が混在しているためである。
  $application_com = fn_Shorten_Com_Type_Name $_applicant.'電力申請会社名称' 
  $manegement_com = fn_Shorten_Com_Type_Name $_applicant.'管理会社名称'
  $employment_com = fn_Shorten_Com_Type_Name $_applicant.'雇用名称'

  if ($manegement_com.Contains("派遣")) { return $manegement_com }
  if ($application_com -eq $employment_com) { return $application_com }
  return ($application_com + "／" + $employment_com)
}

function script:fn_Combined_Name {
  Param(
    [Parameter(Mandatory = $true, Position = 0)][String]$first_name,
    [Parameter(Mandatory = $true, Position = 1)][String]$last_name,
    [String]$delimiter = '　'#デフォルト引数 呼び出し側で -delimiter を指定すること
  )
  $sb = New-Object System.Text.StringBuilder
  #副作用処理  StringBuilderならちょっと速いらしい。要素数が少ないから意味ないかも。
  @($first_name, $delimiter , $last_name) | ForEach-Object { [void] $sb.Append($_) }
  return $sb.ToString()
}



<#  -------- ここから下は実行内容 ----------   #>

Write-Host "登録予約済み情報の出力"
$script:config = fn_Read ".\config\bind_c.json"

$script:regists_Path = (${HOME} + $config.regists_Path)
[PSCustomObject[]]$script:regists_obj_list = fn_Read $regists_Path
#$regists_obj_list.Length

$script:reserved_Path = (${HOME} + $config.reserved_Path)
[PSCustomObject[]]$script:reserved_obj_list = fn_Read $reserved_Path
#$reserved_obj_list | Format-Table
#$reserved_obj_list.Length

$script:new_keys = $config.new_keys

$script:extracted_regists = foreach ($_reserved in $reserved_obj_list) {
  if ($null -eq $_reserved."中央登録番号") { continue }
  [PSCustomObject[]]$filtered_regists = fn_Array_Filter $regists_obj_list "中登番号" $_reserved."中央登録番号"
  # 必要最低限度の予約情報だけ抽出
  $addition_reserved_obj = $_reserved | Select-Object -Property $new_keys
  $extracted = foreach ($_regist in $filtered_regists) {
    fn_Append_KV $_regist $new_keys $addition_reserved_obj
  }
  $extracted  
}
#$extracted_regists | Format-Table
#$extracted_regists.length


# TODO: 副作用 どうにかしたい
# 元のプロパティに再代入するのではなく、新たなプロパティに必要な値を代入する。
foreach ($_ in $extracted_regists) {
  # 登録者管理リストから情報を収集するため、漢字氏名を生成する必要がある。
  $KANJI_name = fn_Combined_Name $_."氏名（姓）" $_."氏名（名）"
  Add-Member -InputObject $_ -NotePropertyName $config.prop_KANJI_name -NotePropertyValue $KANJI_name -Force
  # 新しいプロパティ '解除WBC_申請会社' の代入を設定
  $company_name = fn_WBC_Company_Names ([ref]$_)
  Add-Member -InputObject $_ -NotePropertyName $config.prop_WBC -NotePropertyValue $company_name -Force

  # TODO: 主管課グループ名をどこから参照するか？
  $group_name = "WIDを指定して正式なグループ名を取得したい。"
  Add-Member -InputObject $_ -NotePropertyName "担当主管課班名" -NotePropertyValue $group_name -Force

  # cpy.html で参照しやすいように漢字氏名（姓）と（名）を追加する。
  Add-Member -InputObject $_ -NotePropertyName "漢字氏名（姓）" -NotePropertyValue $_."氏名（姓）" -Force
  Add-Member -InputObject $_ -NotePropertyName "漢字氏名（名）" -NotePropertyValue $_."氏名（名）" -Force
  Add-Member -InputObject $_ -NotePropertyName "氏名（カナ）" -NotePropertyValue (fn_To_Wide $_."氏名（カナ）") -Force

  $original_company_number = ($_."電力申請会社番号").replace("T", "0")
  Add-Member -InputObject $_ -NotePropertyName "電力申請会社番号" -NotePropertyValue $original_company_number -Force
}

$selection = $extracted_regists | Select-Object -Property $config.c_selection
$selection | Format-Table

$utf8_BOM = New-Object System.Text.UTF8Encoding $True
$json_path = (${HOME} + $config.export_Path)
fn_Write_JSON $json_path $selection $utf8_BOM
