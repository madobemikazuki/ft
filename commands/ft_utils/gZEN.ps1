Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"
<#
  \Downloads\PAN配下の *事前申請*.txt を読み込んで
  \config\csv_header\gZEN_header_ANSI.txtとともに
  申請会社情報 : Downloads\TEMP\Companies.json
  を中央登録番号に基づきに格納し、
  \config\csv_header\gZEN_sorted_ANSI.txt に指定したフィールドの順番で
  JSON形式を出力する。
#>

. ..\ft_cores\FT_IO.ps1
. ..\ft_cores\FT_Name.ps1
. ..\ft_cores\FT_Dict.ps1
. ..\ft_cores\FT_Array.ps1
. ..\ft_cores\FT_Log.ps1
. ..\ft_cores\FT_Message.ps1

function script:fn_Post_Code {
  Param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern("^\d{7}")][String]$arg
  )
  return Write-Output("{0:000-0000}" -f [Int]$arg)
}


function script:fn_To_Wide {
  Param(
    [Parameter(Mandatory = $True)][String]$half_string
  )
  Add-Type -AssemblyName "Microsoft.VisualBasic"
  [Microsoft.VisualBasic.Strings]::StrConv($half_string, [Microsoft.VisualBasic.VbStrConv]::Wide)
}

function script:fn_Sort_by_Uniqued_Array {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject[]]$_arr,
    [Parameter(Mandatory = $True, Position = 1)]
    [String[]]$_selection_field_arr
  )
  # オブジェクトのプロパティ名をプライマリキーとしてソートして重複を排除する
  $uniqued_arr = [FT_Array]::SortByUnique($_arr, "中央登録番号")
  $sorted_arr = [FT_Array]::Map($uniqued_arr, $_selection_field_arr)
  remove-variable uniqued_arr
  return $sorted_arr
}


function fn_Shape_Values {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [HashTable]$_hash,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject]$_config,
    [Parameter(Mandatory = $True, Position = 2)]
    [PSCustomObject]$_addition
  )
  [PSCustomObject]$private:coms = [FT_IO]::Read_JSON_Object((${HOME} + $_config.app_coms_path))
  [PSCustomObject]$app_coms = $coms.app_coms

  $coms_path = (${HOME} + $_config.app_coms_path) 
  #$addition = $_config.addition_field
  $source_field = $_config.source_field
  $delimiter = $_config.name_delimiter

  [PSCustomObject[]]$new_obj = foreach ($_key in $_hash.keys) {
    $obj = $_hash.$_key
    Add-Member -InputObject $obj -NotePropertyMembers @{
      $_addition.FT_name_kanji = [FT_Name]::Binding(
        $obj.($source_field.second_name_kanji),
        $obj.($source_field.first_name_kanji),
        $delimiter);
      $_addition.FT_name_kana  = [FT_Name]::Binding(
        $obj.($source_field.second_name_kana),
        $obj.($source_field.first_name_kana),
        $delimiter);
    }

    #既存の値を変更
    $obj.($source_field.current_zip_code) = fn_Post_Code $obj.($source_field.current_zip_code)
    $obj.($source_field.current_address) = fn_To_Wide $obj.($source_field.current_address)
    $obj.($source_field.company_number) = $obj.($source_field.company_number) -replace "^0", "T"

    # 所属企業名の割り当て
    # この処理は 各オブジェクトに所属企業番号 の正規化した値が存在しなければ機能しない。
    if ($app_coms.psobject.properties.name -notcontains ($obj.($source_field.company_number))) {
      $cmd_name = Split-Path -Leaf $PSCommandPath
      $log_type = "ループ中のスキップ処理"
      $message = "指定した所属企業番号の値が存在しないため処理をスキップしました。"
      $purpose = "本来はスキップしてはならない処理がスキップされたのでログに記録する。"
      $log_hash = [FT_Log]::Create($cmd_name, $log_type, $obj.'所属企業番号', $message, $purpose)
      Write-Host $obj.'所属企業番号'$message
      [FT_Log]::Write(${HOME} + $_config.log_path, $log_hash)
      Start-Process -FilePath notepad.exe -ArgumentList $coms_path
      #最後のループ要素を continue するとログの書き込み処理が実行できない。
      continue
    }
    Add-Member -InputObject $obj -NotePropertyMembers @{
      $_addition.FT_company_name = $app_coms.($obj.'所属企業番号')
    }
    $obj
  }
  return $new_obj
}


<#
--------------------- ここから実行内容 -----------------------
#>
$private:config_path = ".\config\FT_Utils.json"
$private:commandlet_name = Split-Path -Leaf $PSCommandPath

[PSCustomObject]$private:config = [FT_IO]::Read_JSON_Object($config_path)
$private:gZEN_config = $config.$commandlet_name
$private:addition = $config.common_field

[FT_Message]::execution($gZEN_config.command_name)
remove-variable config

$txt_encoding = "Default"
$header = [FT_IO]::Read_ToArray($gZEN_config.orign_header, $txt_encoding)

# *事前申請*.txt はSHIFT-JISでエンコードされている。
$pan_folder = (${HOME} + $gZEN_config.temp_folder)
$values_filenames = [FT_IO]::Find($pan_folder, $gZEN_config.gZEN_targets)
#$values_filenames

$values = foreach ($_ in $values_filenames) {
  [FT_IO]::Read_ToArray($_, $txt_encoding)
  # 使用済みの事前承認ファイルを waste フォルダへ移動する?
  #Move-Item -Path $_ -Destination ($pan_folder + $config.waste_folder)
}

#テキストファイルを読み込み、ヘッダーをつけてCSVファイルを生成。
[PSCustomObject[]]$csv_obj = [FT_IO]::Bind_As_CSV($header, $values, ',')
$hash = [FT_Array]::ToDict($csv_obj, $gZEN_config.primary_key)
$selection_field = [FT_IO]::Read_ToArray($gZEN_config.sorted_header, $txt_encoding) 

remove-variable -Name config_path, header, values, values_filenames, pan_folder

# TODO: リファクタリングすること
# 必要な項目の情報を抽出する
$shaped_obj = fn_Shape_Values $hash $gZEN_config $addition


# 出力プロパティを任意の順序で取得する。
$final_arr = [FT_Array]::Map($shaped_obj, $selection_field)
remove-variable shaped_obj


$export_path = (${HOME} + $gZEN_config.export_JSON_path)
[FT_IO]::Write_JSON_Array($export_path, $final_arr)

#clipboardに出力する
<#
$private:comma = ', '
$private:tab = "`t"
$plain_text = Get-Content -Path $CSV_destination -Encoding UTF8
$formatted_text = $plain_text.Replace('"', '')
$formatted_text.Replace($comma, $tab) | Set-Clipboard
#>

#ftの親パスを取得
$private:app_path = Convert-Path .
$private:cpy_path = ($app_path + $gZEN_config.cpy_monolith_path)
# EDGEブラウザ起動し、cpyを起動する。
Start-Process msedge $cpy_path

#Write-Host ('🐈., 💩💩💩  ' + $command_name + ' ' + $config.version + "::出力完了 => " + $JSON_destination)
exit 0

