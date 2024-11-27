Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [ValidatePattern("ts|tp|ih|TS|TP|IH")]
  [String]$_Tgroup,
  [Parameter(Mandatory = $True, Position = 1)]
  [ValidatePattern("c|d|cd|j|C|D|CD|J")]
  [String]$_Kinds,
  [Parameter(Mandatory = $true, Position = 2)]
  [ValidatePattern("^\d{2}\b\-\b\d{6}$")]
  [String[]]$_central_nums
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

. .\ft_cores\FT_IO.ps1
. .\ft_cores\FT_Dict.ps1
. .\ft_cores\FT_Name.ps1
. .\ft_cores\FT_Array.ps1
. .\ft_cores\Poe\PoeObject.ps1
. .\ft_cores\Poe\PoeAddress.ps1


<#
1.事前申請者の登録予約情報から中央登録番号の該当者を検索する
2.既存の登録者の中から中登番号の該当者を検索する
上記、いずれかの該当者の情報を参照し、教育実施記録を出力する。
#>
function fn_Generate_Export_Path {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [ValidatePattern("ts|tp|ih|TS|TP|IH")]
    [String][ref]$_Tgroup,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject][ref]$_config,
    [Parameter(Mandatory = $True, Position = 2)]
    [String][ref]$_applicants_names
  )
  $export_name = @(
    "${HOME}",
    $_config.$_Tgroup.export_folder,
    $_config.File.head_name,
    $_applicants_names,
    $_config.File.extension
  ) -join ""
  return  [FT_Name]::One_Liner($export_name)
}


[PSCustomObject]$config = [FT_IO]::Read_JSON_Object(".\config\ed.json")
#$config.TS.export_folder 
#Test-Path (${HOME} + $config.gZEN_csv)

$reserved_source_Path = (${HOME} + $config.reserved_info.path)
[PSCustomObject[]]$reserved_arr = [FT_IO]::Read_JSON_Array($reserved_source_Path)

$primary_key = $config.primary_key
$reserved_dict = [FT_Array]::ToDict($reserved_arr, $primary_key)

$header = $config.extraction
# 申請用オブジェクト生成 --------------------------------------


# 登録予約者情報に該当者が存在する場合 not equals
if ([FT_Dict]::Every($reserved_dict, $_central_nums)) {
  $targets = [FT_Dict]::Search($reserved_dict, $_central_nums)
  $script:applicants_dict = [FT_Dict]::Selective($targets, $header)
  Write-Host "登録予約者のなかにおったよ。"
  Remove-Variable reserved_source_Path, reserved_arr, targets
}

# 登録予約者に該当者が存在しない場合 既存の登録者の中から探す
if (![FT_Dict]::Every($reserved_dict, $_central_nums)) {
  Write-Host "予約情報のなかにはおらんやったよ。"

  $private:registerers_path = (${Home} + $config.registerers_info.path)
  $private:registerers_arr = [FT_IO]::Read_JSON_Array($registerers_path)
  $private:registeres_dict = [FT_Array]::ToDict($registerers_arr, $primary_key)
  if ([FT_Dict]::Every($registeres_dict, $_central_nums)) {
    $private:targets = [FT_Dict]::Search($registeres_dict, $_central_nums)
    $script:applicants_dict = [FT_Dict]::Selective($targets, $header)
    Write-Host "既登録者のなかにおったよ。"
    #$applicants_dict.Values | Format-List
  }
  else {
    Write-Host "既登録者のなかにもおらんかったよ。"
    exit 404
  }
  #無限ループでアプリを使い回すなら、変数を消すと速度低下になるか？
  Remove-Variable primary_key, registerers_path, registerers_arr, registeres_dict
}

# まさか登録予定者と定期受検者を混在させて入力することはないだろう。
#$applicants_dict.Values | Format-Table



# TODO: 名前が適当ではない
# このオブジェクトを使い回す
$private:subject = @{
  Tgroup   = $_Tgroup;
  kinds    = $_Kinds;
  template = (${HOME} + $config.template_path);
}
$name_field = $config.File.applicant
foreach ($_applicant in $applicants_dict.Values) {
  # applicant ごとにエクスポートファイル名は変わる
  $export_path = fn_Generate_Export_Path ([ref]$_Tgroup) ([ref]$config) ([ref]$_applicant.$name_field);
  . .\ft_cores\Poe\Poe-TranscriptionByKinds.ps1 $_applicant $config $subject $export_path
}

exit 0

