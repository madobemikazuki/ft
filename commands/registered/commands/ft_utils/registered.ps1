Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

. ..\ft_cores\FT_IO.ps1
. ..\ft_cores\FT_Name.ps1
. ..\ft_cores\FT_Dict.ps1
. ..\ft_cores\Odd_Name.ps1
. ..\ft_cores\FT_Specific_Funcs.ps1

# 登録情報から必要最低限の情報のみを抽出して json ファイルに出力する。

$config_path = ".\config\FT_Utils.json"
$command_name = Split-Path -Leaf $PSCommandPath
$config = ([FT_IO]::Read_JSON_Object($config_path))
$rgst_config = $config.$command_name
$additions = $config.common_field
remove-variable config

$csv_object_arr = [FT_IO]::Read_CSV((${HOME} + $rgst_config.import_csv_path))
$primary_key = $rgst_config.primary_key

# hashtable の中身は  [String]key = [PSCustomObject]value となっている。
$dict = [FT_Dict]::Convert($csv_object_arr, $primary_key)
Remove-Variable csv_object_arr
$registerer_dict = [FT_Dict]::Selective($dict, $rgst_config.first_extraction_target)


$convs = $rgst_config.convs
$delimiter = '　'
$odd_names_path = (${HOME}+$rgst_config.odd_dict_path)
$odd = [FT_IO]::Read_JSON_Object($odd_names_path)

$new_dict = @{}
Add-Type -AssemblyName "Microsoft.VisualBasic"
#$new_arr = 
foreach ($_ in $registerer_dict.keys) {
  $obj = $registerer_dict.$_
  Add-Member -InputObject $registerer_dict.$_ -NotePropertyMembers @{
    $additions.FT_name_kanji = [FT_Name]::Binding($obj.($convs.second_name), $obj.($convs.first_name), $delimiter);
    $additions.FT_name_kana  = [Microsoft.VisualBasic.Strings]::StrConv($obj.($convs.name_kana), [Microsoft.VisualBasic.VbStrConv]::Wide);
    $additions.FT_app_company_number = [FT_Name]::Replace_Head($obj.($convs.company_number), 'T', '0');
    $additions.FT_app_company_name = [FT_Name]::Shortened_Com_Type_Name($obj.($convs.company_name));
    $additions.FT_ed_company_names = [Odd_Name]::To_Appropriate((fn_SF_C_Company_Names $obj), $odd);
    $additions.central_number = $_
  }
  $new_dict[$_] = $obj
  #$obj
}

Remove-Variable registerer_dict

$final_dict = [FT_Dict]::Selective($new_dict, $rgst_config.final_extraction_target)

$export_path = $rgst_config.export_json_path
#[FT_IO]::Write_JSON_Array((${HOME} + $export_path), $new_arr)
$final_arr = foreach($_ in $final_dict.keys){
  $final_dict.$_
}
[FT_IO]::Write_JSON_Array((${HOME} + $export_path), $final_arr)

