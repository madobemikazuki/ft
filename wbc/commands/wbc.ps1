Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [ValidatePattern("r|c")]$Task,
  [Parameter(Mandatory = $True, Position = 2)]
  [ValidatePattern("^\d{2}\b\-\b\d{6}$")]
  [String[]]$_central_nums
)

<#
  動作チェック
  .\commands\wbc.ps1 r のときは
  \Downloads\TEMP\gZEN_exported.csv から "中央登録番号" を選択することになる

  .\commands\wbc.ps1 c のときは
  \Downloads\TEMP\登録者管理リスト.csv から "中登番号"を選択することになる
#>

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"



function fn_Excel_Hell_Format {
  Param(
    [Parameter(Mandatory)]
    [DateTime]$_future
  )
  return $_future.ToString("yyyy年　MM月　dd日");
}


function script:fn_Shorten_Com_Type_Name {
  Param(
    [Parameter(Mandatory = $True)]
    [String]$_corporate_name
  )
  switch ($_corporate_name) {
    { $_corporate_name.Contains('株式会社') } { return $_corporate_name.Replace('株式会社', '（株）') }
    { $_corporate_name.Contains('有限会社') } { return $_corporate_name.Replace('有限会社', '（有）') }
    #半角カッコを全角カッコにする
    { $_corporate_name.Contains('(株)') } { return $_corporate_name.Replace('(株)', '（株）') }
    { $_corporate_name.Contains('(有)') } { return $_corporate_name.Replace('(有)', '（有）') }
    default { return $_corporate_name }
  }
}

function script:fn_Search_Target {
  Param(
    [Parameter(Mandatory = $True, Position = 0) ]
    [PSCustomObject[]][ref]$_PSCO_array,
    [Parameter(Mandatory = $True, Position = 1) ]
    [String[]]$_reg_num_list,
    [Parameter(Mandatory = $True, Position = 2)]
    [String]$_flag
  )
  # $_flag には'中央登録番号' もしくは '中登番号' が入る
  [PSCustomObject[]]$result = foreach ($_ in $_PSCO_array) {
    if ($_reg_num_list.Contains($_.$_flag)) {
      $_
    }
  }
  return $result
}

function fn_PSCO_List_Filter {
  Param(
    [Parameter(Mandatory = $True, Position = 0) ]
    [PSCustomObject[]][ref]$_object_array,
    [Parameter(Mandatory = $True, Position = 1) ]
    [String[]]$_targets
  )
  $collection = foreach ($object in ($_object_array)) {
    $object | Select-Object -Property $_targets
  }
  return $collection
}

function script:fn_Read {
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


function fn_Generate_WBC_Export_Path {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [ValidatePattern('r|c')]$_Task,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject][ref]$_wbc_config,
    [Parameter(Mandatory = $True, Position = 2)]
    [String]$_names
  )

  $output_file_path = @(
    $_wbc_config.$_Task.output_folder,
    $_wbc_config.$_Task.task,
    'WBC',
    '_',
    $_names,
    $_wbc_config.extension
  ) -Join ''
  return (${HOME} + $output_file_path)
}

# TEPCOに提出した登録の事前申請書 から情報を取得し、整形して返す
function fn_Registration_Format {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String[]]$_central_num_list,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject]$_config_r,
    [Parameter(Mandatory = $True, Position = 2)]
    [String[]]$_printig_field
  )
  Write-Host ($_config_r.task + 'モードだよ')

  #登録_予約済み申請者リスト_UTF8-bom.json から情報を取得
  $reserved_list = fn_Read (${HOME} + $_config_r.source_path)
  
  $search_flag = '中央登録番号'
  #事前申請書のヘッダーに基づく抽出対象
  $reserved_app_list = fn_Search_Target ([ref]$reserved_list) $_central_num_list $search_flag
  [PSCustomObject[]]$private:extracted_targets_info = fn_PSCO_List_Filter ([ref]$reserved_app_list) ($_config_r.extraction_list)

  [PSCustomObject[]]$private:applicant_list = foreach ($_ in $extracted_targets_info) {
    $private:field = $_printig_field
    [PSCustomObject] @{
      $field[0] = fn_Excel_Hell_Format ([DateTime]$_.'登録予約日')
      $field[1] = $_.'個人番号'
      $field[2] = fn_Shorten_Com_Type_Name $_.'登録_申請会社'
      $field[3] = $_.'漢字氏名'
    }
  }
  return $applicant_list
}

function fn_Cancellation_Format {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [ValidatePattern("^\d{2}\b\-\b\d{6}$")]
    [String[]]$_central_num_list,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject]$_config_c,
    [Parameter(Mandatory = $True, Position = 2)]
    [String[]]$_printig_field
  )
  Write-Host ($_config_c.task + 'モードだよ')
  $private:reserved_list = fn_Read (${HOME} + $_config_c.source_path)

  $private:search_flag = '中登番号'
  $private:reserved_app_list = fn_Search_Target ([ref]$reserved_list) $_central_num_list $search_flag
  # 登録者管理リスト.xls のヘッダーに基づく抽出対象
  [PSCustomObject[]]$private:extracted_targets_info = fn_PSCO_List_Filter ([ref]$reserved_app_list) ($_config_c.extraction_list)
  foreach ($_ in $extracted_targets_info) {
    $private:field = $_printig_field
    [PSCustomObject] @{
      # TODO:wbc_application_field を外部スコープに依存しているのがきもいなぁ。
      $field[0] = fn_Excel_Hell_Format ([DateTime]$_."解除予約日")
      $field[1] = $_."作業者証番号"
      
      # ◆未実装  所属企業番号から '申請会社名' を取得すること 
      $field[2] = fn_Shorten_Com_Type_Name $_.'解除WBC_申請会社'
      $field[3] = $_.'漢字氏名'
    }
  }
}

function private:fn_display {
  Param(
    [Parameter(Mandatory = $True)]
    [String]$_message
  )
  return  $_message + '楽しんでね！'
}


function fn_PSObjList_Map {
  Param(
    [Parameter(Mandatory = $True, Position = 0) ]
    [PSCustomObject[]][ref]$_object_array,
    [Parameter(Mandatory = $True, Position = 1) ]
    [String[]]$_targets
  )
  $collection = foreach ($object in ($_object_array)) {
    $object | Select-Object -Property $_targets
  }
  return $collection
}

function fn_Unit_Processing {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject[]]$_applicants,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject]$_config_object,
    [Parameter(Mandatory = $True, Position = 2)]
    [String]$_Task
  )
  $field = $_config_object.poe_config.printing.printig_field
  #氏名をピックアップしてファイル名に利用する。
  $name_list = fn_PSObjList_Map ([ref]$_applicants) @($field[3])
  $names = foreach ($name in $name_list) { $name.($field[3]) }
  $one_line_names = $names -join '_'
  $one_line = $one_line_names.replace("　", "")
  $export_path = fn_Generate_WBC_Export_Path $_Task ([ref]$_config_object) $one_line
  #$export_path
  $address_table = $_config_object.poe_config.address_table
  $posting_object = [PoeAddress]::Unit_Format($_applicants, $address_table)
  $posting_object | Format-Table

  # $export_path に出力
  . .\ft_core\Poe\Poe-Transcription.ps1 $posting_object $_config_object.poe_config $export_path
}

<#
------------------- ここからコマンドの実行内容   ---------------------------------
#>
. .\ft_core\Poe\PoeObject.ps1
. .\ft_core\Poe\PoeAddress.ps1

$private:config_path = ".\config\wbc.json"
[PSCustomObject]$script:config_object = fn_Read $config_path


#TODO: $config_object.source_path から直接中央登録番号の配列を取得すれば楽じゃなか？
# それぞれのフォーマット出力関数内で抽出すればいけるな。
#ここで必要な情報を集約整形して格納する。
$printing_field = $config_object.poe_config.printing.printig_field
Write-Host "Hello"
switch ($Task) {
  r { $script:applicants = fn_Registration_Format $_central_nums $config_object.$Task $printing_field}
  c { $script:applicants = fn_Cancellation_Format $_central_nums $config_object.$Task $printing_field}
}
$applicants | Format-Table


$max_range = $config_object.poe_config.max_range
#TODO:  以下の処理を Poe.ps1 に集約したい. vvv このようにしたい。
#TODO:  . .\ft_core\Poe2\Poe.ps1 $config $applicants $export_path
# $applicants.length が1より大きく5より小さければ fn_Unit_Processing を一回実行
if (
    (0 -lt $applicants.length) -and
    ($applicants.length -lt ($max_range + 1))) {
  Write-Host "単一処理だよ"
  fn_Unit_Processing $applicants $config_object $Task
  exit 0
}

# $applicants.length が 4より大きければチャンクする
if ($max_range -lt $applicants.length) {
  #$applicants | Format-Table
  Write-Host "チャンク処理だよ"
  $jugged_arr = . .\ft_core\list\Jugged-Array.ps1 $applicants $max_range
  Write-Host $jugged_arr.length " : ジャグ配列の数"
  foreach ($_chunk in $jugged_arr){
    fn_Unit_Processing $_chunk $config_object $Task
  }
  exit 0
}

