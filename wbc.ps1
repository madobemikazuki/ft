Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [ValidatePattern("r|c")]$_Task,
  [Parameter(Mandatory = $True, Position = 2)]
  [ValidatePattern("^\d{2}\b\-\b\d{6}$")]
  [String[]]$_central_nums
)

<#
  動作チェック
  .\commands\wbc.ps1 r のときは
  \Downloads\TEMP\gZEN_exported.json から "中央登録番号" を選択することになる

  .\commands\wbc.ps1 c のときは
  \Downloads\PAN\Registered_UTF8-bom.json から "中登番号"を選択することになる
#>

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

. .\ft_cores\FT_IO.ps1
. .\ft_cores\FT_Name.ps1
. .\ft_cores\FT_Dict.ps1
. .\ft_cores\FT_Array.ps1
. .\ft_cores\Poe\PoeObject.ps1
. .\ft_cores\Poe\PoeAddress.ps1

# ---------------------------------------------------------
# 関数群
function fn_Generate_WBC_Export_Path {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [ValidatePattern('r|c')]$_Task,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject][ref]$_wbc_config,
    [Parameter(Mandatory = $True, Position = 2)]
    [String]$_names
  )
  $private:delimiter = '_'
  $private:output_file_path = @(
    $_wbc_config.$_Task.output_folder,
    $_wbc_config.$_Task.task,
    $_wbc_config.file_label,
    $delimiter,
    $_names,
    $_wbc_config.extension
  ) -Join ''
  return (${HOME} + $output_file_path)
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
  $private:_name_field = ($_config_object.poe_config.printing.printig_field)[3]
  #氏名をピックアップしてファイル名に利用する。
  $private:one_line = [FT_Name]::One_Liner($_applicants, $_name_field)
  #Write-Host $one_line
  #Write-Host $export_path
  
  $address_table = $_config_object.poe_config.address_table
  $posting_object = [PoeAddress]::Unit_Format($_applicants, $address_table)
  #$posting_object | Format-Table

  
  $private:paths = @{
    template = (${Home} + ($_config_object.$_Task).tamplate_path);
    export   = fn_Generate_WBC_Export_Path $_Task ([ref]$_config_object) $one_line
  }
  
  
  # $export_path に出力
  . .\ft_cores\Poe\Poe-Transcription.ps1 $posting_object $_config_object.poe_config $paths
}

# TEPCOに提出した登録の事前申請書 から情報を取得し、整形して返す
function fn_WBC_Format {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String[]]$_central_num_list,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject]$_config,
    [Parameter(Mandatory = $True, Position = 2)]
    [String[]]$_printing_field
  )
  Write-Host ($_config.task + 'モードです。')

  $primary_key = $_config.primary_key
  #登録_予約済み申請者リスト_UTF8-bom.json から情報を取得
  $private:reserved_list = [FT_IO]::Read_JSON_Array(${HOME} + $_config.source_path)
  $private:reserved_dict = [FT_Array]::ToDict($reserved_list, $primary_key)
  
  # 該当しない人がいる場合は処理を中止する。
  # TODO: 処理の共通化のためにもっと簡潔にできないだろうか。

  $private:target_dict = [FT_Dict]::Search($reserved_dict, $_central_num_list)
  $private:extracted_dict = [FT_Dict]::Selective($target_dict, $_config.extraction_list)
  #$extracted_dict | Format-List
  $private:formated_arr = foreach ($_key in $extracted_dict.Keys) {
    $private:_obj = $extracted_dict.$_key
    [PSCustomObject] @{
      $_printing_field[0] = $_obj.($_config.extraction_list[1]);
      $_printing_field[1] = $_obj.($_config.extraction_list[2]);
      $_printing_field[2] = $_obj.($_config.extraction_list[3]);
      $_printing_field[3] = $_obj.($_config.extraction_list[4]);
    }
  }
  Remove-Variable primary_key, reserved_list, reserved_dict
  return $formated_arr
}

# -----------------------------------------------------------------
$private:config_path = ".\config\wbc.json"
[PSCustomObject]$script:config_object = [FT_IO]::Read_JSON_Object($config_path)

$printing_field = $config_object.poe_config.printing.printig_field
[PSCustomObject[]]$applicants = fn_WBC_Format $_central_nums $config_object.$_Task $printing_field
#$applicants | Format-Table

$max_range = $config_object.poe_config.max_range
#$max_range

if ($null -eq $applicants) {
  Write-Host "該当者がいませんでした。処理を停止します。" -ForegroundColor RED
  exit 404
}

if ( (0 -lt $applicants.Length) -and ($applicants.Length -lt ($max_range + 1))) {
  Write-Host "単一処理の開始"
  fn_Unit_Processing $applicants $config_object $_Task
}

# $applicants.length が 4より大きければチャンクする
if ($max_range -lt $applicants.length) {
  Write-Host "チャンク処理の開始"
  $jugged_arr = [FT_Array]::Jugged($applicants, $max_range)
  #Write-Host $jugged_arr.length " : ジャグ配列の数"
  foreach ($_chunk in $jugged_arr) {
    fn_Unit_Processing $_chunk $config_object $_Task
  }
}

