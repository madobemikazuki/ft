Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [ValidatePattern("^\d{8}")]
  [String]$_date,
  [Parameter(Mandatory = $True, Position = 1)]
  [ValidatePattern("^\d{6}")]
  [String]$_wid,
  [Parameter(Mandatory = $True, Position = 2)]
  [ValidatePattern("^\d{2}\b\-\b\d{6}$")]
  [ValidateCount(1, 10)]
  [String[]]$_central_nums
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"


function script:fn_DisQualification {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [hashtable][ref]$_dict,
    [Parameter(Mandatory = $True, Position = 1)]
    [String[]][ref]$_prequisites
  )
  # 必須条件としての値が一意であること。
  $boolean_list = foreach ($_preq in $_prequisites) {
    # @を省くと自動的に単一のObject が返ってくる。
    @([FT_Dict]::Selective($_dict, $_prequisites) | Sort-Object -Property $_preq -Unique).length
  }
  $boolean_list | Format-List
  $result = foreach ($_ in $boolean_list) { $_ -eq 1 }
  Write-Host $result
  return ($result -contains $False)
}



# wid_group の入力を催促する
function fn_Urge {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$wid_group_path,
    [Paramater(Mandatory = $True, Position = 1)]
    [String]$_wid_num
  )
  Start-Process notepad.exe $wid_group_path
  Set-Clipboard $_wid_num
}

function fn_Generate_Export_Path {
  Param(
    [Parameter(Mandatory = $True, Position = 0)][PSCustomObject]$_export_config,
    [Parameter(Mandatory = $True, Position = 1)][String[]]$_names
  )  
  $private:folder = (${HOME} + $_export_config.folder)
  $private:head_name = $_export_config.file_name.first
  $private:names = $_names -join $_export_config.file_name.conjunction
  $private:shorten_names = $names -replace $_export_config.file_name.replaces
  $private:extension = $_export_config.file_name.extension
  return ($folder + $head_name + $shorten_names + $extension)
}

#--------------------------------------------------------------------------------
. .\ft_cores\FT_IO.ps1
. .\ft_cores\FT_Date.ps1
. .\ft_cores\FT_Dict.ps1
. .\ft_cores\FT_Array.ps1
. .\ft_cores\Poe\PoeObject.ps1
. .\ft_cores\Poe\PoeAddress.ps1

$config = [FT_IO]::Read_JSON_Object(".\config\cnc.json")

$wid_path = (${HOME} + $config.wid_path)
$wid_LookUpHash = [FT_IO]::Read_JSON_Object($wid_path)
try { 
  $script:wid = $wid_LookUpHash.$_wid
  if ([string]::IsNullOrEmpty($wid.group)) {
    fn_Urge $wid_path $_wid
    Throw ("💩 WID : " + $_wid + " の group (担当主管課班名) が存在しないので追記してください。💩")
  }
}
catch {
  fn_Urge $wid_path $_wid
  Throw ("💩 Error : 指定した WID: " + $_wid + " は存在しないので追記してください。💩")
}
#Write-Host $wid.depertment
#Write-Host $wid.group

# remove-variable の変数名に $ は不要です。
remove-variable wid_LookUpHash
$script:poe_config = $config.poe_config


# 該当者を検索

# 予約者情報
$reserved_arr = [FT_IO]::Read_JSON_Array(${HOME} + $config.reserved_path)
$reserved_dict = [FT_Array]::ToDict($reserved_arr, $config.primary_key)

# 既存の登録者情報
$registered_arr = [FT_IO]::Read_JSON_Array(${HOME} + $config.registerer_path)
$registered_dict = [FT_Array]::ToDict($registered_arr, $config.primary_key)

$result = [PScustomObject[]]@()
if ([FT_Dict]::Every($reserved_dict, $_central_nums)) {
  Remove-Variable registered_arr, registered_dict
  # 共通利用の関数orメソッド化
  $target_dict = [FT_Dict]::Search($reserved_dict, $_central_nums)
  $extracted_dict = [FT_Dict]::Selective($target_dict, $config.extraction_list)
  #$extracted_dict.Values | Format-List
  $result = $extracted_dict.Values
  #$result |Format-Table

  Write-Host "予約者情報を参照した。"
}
elseif ([FT_Dict]::Every($registered_dict, $_central_nums)) {
  Remove-Variable reserved_arr, reserved_dict
  # 共通利用の関数orメソッド化
  $target_dict = [FT_Dict]::Search($registered_dict, $_central_nums)
  $extracted_dict = [FT_Dict]::Selective($target_dict, $config.extraction_list)
  #$extracted_dict.Values | Format-List
  $result = $extracted_dict.Values
  Write-Host "既存の登録者情報を参照した。"
}
else {
  Write-Host "該当者はいませんでした。"
  exit 404
}

<#
# TODO:
if (($target_dict.count -gt 1) -and (fn_DisQualification ([ref]$target_dict) ([ref]$config.prequisites))) {
  throw "申請者全員の申請会社名称、もしくは雇用名称が異なりますね。💩"
  exit 0
}
#>
$temp_name = ($wid.depertment + "`r`n" + $wid.group) -replace '[ＧG]$', 'グループ'
Add-Type -AssemblyName "Microsoft.VisualBasic"  
$common_obj = [PSCustomObject]@{
  # フィールド名を config に切り出す
  "解除予約日"   = [FT_Date]::Slash_Format($_date);
  "担当主管課班名" = [Microsoft.VisualBasic.Strings]::StrConv($temp_name, [Microsoft.VisualBasic.VbStrConv]::Wide)
}
# $common_obj | Format-List


$common_address = $poe_config.common_address_table
$common_obj_list = [PoeAddress]::Common_Format($common_obj, $common_address)

$main_address_table = $poe_config.address_table
$main_obj_list = [PoeAddress]::List_Format($result, $main_address_table)

# 転記情報を統合する
$poe_obj_list = $common_obj_list + $main_obj_list
$poe_obj_list | Format-Table

$names = [FT_Array]::Map($result, @('FT_氏名_漢字'))
$export_path = fn_Generate_Export_Path $poe_config.export $names

# 最終的な出力を行う
. .\ft_cores\Poe\Poe-Transcription.ps1 $poe_obj_list $poe_config $export_path

