Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

<#
  オブジェクト形式で保存  \TEMP\WID_LookUpHash.json
  配列形式で保存  \PAN\WID_min_UTF8-bom.json

  高速化するには配列の使用を止めるべきだが、
  リファクタリングする暇がない。
#>

. ..\ft_cores\FT_IO.ps1
. ..\ft_cores\FT_Array.ps1


function fn_PSCOjb_Arr {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject[]]$_Object_List
  )
  $private:dict = @{}
  foreach ($_ in $_Object_List) {
    $private:new_obj = [ordered]@{}
    $new_obj["wid"] = $_."作業件名コード"
    $new_obj["subject"] = $_."作業件名"
    $new_obj["depertment"], $new_obj["group"] = $_."作業主管グループ".Split("　")
    $new_obj["document_path"] = ""
    $dict[$new_obj.wid] = $new_obj
  }
  return $dict.Values
}
# ------------------------------------------------------------------------


$config_path = ".\config\FT_Utils.json"
$command_name = Split-Path -Leaf $PSCommandPath

[PSCustomObject]$script:config = ([FT_IO]::Read_JSON_Object($config_path)).$command_name

# 対象フォルダー
$source_folder = (${HOME} + $config.import.folder)
$name = $config.import.contained_name
$script:wid_XLS_file_path = [FT_IO]::Find_Latest_File($source_folder, $name)
Write-Host $wid_XLS_file_path


try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $False
  $excel.DisplayAlerts = $False

  $book = $excel.Workbooks.Open($wid_XLS_file_path, 0, $True)
  $target_page = 1
  $Sheet = $book.Sheets.Item($target_page)

  #行 ( rows y軸) の設定 値取得対象の始点行から最終行の設定
  $starting_row = $config.import.starting_row
  $end_of_rows = $Sheet.UsedRange.Rows.Count + 1
  $select_rows_range = @($starting_row..$end_of_rows)
   
  # 列 ( columns x軸)の設定
  $starting_column = $config.import.starting_column
  $end_of_columns = $config.import.end_of_columns
  $columns = @($starting_column..$end_of_columns)
  $export_field = $config.export.field


  # PSCustomObject[]に格納する。
  $wid_regexp = $config.export.wid_regexp
  [PScustomObject[]]$list = foreach ($_row in $select_rows_range) {
    # pscustomObject に格納する。
    # return object


    # TODO: WID の値が　$config.export.wid_regexp　に該当しないものは continue する。
    # 理由: データの取り込み処理が遅すぎる。古いデータは不要。
    if (($Sheet.Cells.Item($_row, $config.continue_column_number).Text) -notmatch $wid_regexp) { continue }
    $private:obj = @{}
    foreach ($_column in $columns) {
      $index = $columns.IndexOf($_column)
      $key = $export_field[$index]
      $value = $Sheet.Cells.Item($_row, $_column).Text
      $obj[$key] = $value
    }
    [PSCustomObject]$obj
  }
  $script:WID_KEY = $config.export.primary_key
  $new_list = [FT_Array]::SortDesc($list, $WID_KEY)



  # 最小限の情報を JSON に出力する
  $min_selcets = $config.export.min_field
  [PSCustomObject[]]$script:addition_list = [FT_Array]::Map($new_list, $min_selcets) 
  #foreach ($_ in $final_list) {
  #  $_ | Select-Object -Property $min_selcets
  #}

  #Write-Host "addition_list 項目追加されている新しいオブジェクトのリスト"
  #Write-Host "xlsから読み込んだもの"
  #$addition_list | Format-Table
  $exists_json_path = (${HOME} + $config.export.min_json_path)
  $customized_wid_path = (${HOME} + $config.export.customized_wid_path)

  # $addition_list が完成したのち、既存の WID_min_UTF8-bom.json
  if (Test-Path $exists_json_path) {
    #Write-Host "Values"
    $exists_obj_list = [FT_IO]::Read_JSON_Object($exists_json_path)
    #$exists_obj_list | Format-Table



    . ..\ft_cores\ADF.ps1
    [PSCustomObject[]]$script:complete_list = ADF $exists_obj_list $addition_list $WID_KEY



    if (($complete_list[0] -eq 0) -or ($Null -eq $complete_list)) {
      Write-Host $config.command_name "追記するものなし。更新せずに終了しました。"
      exit 0
    }

    $final_list = $complete_list | Sort-Object -Property $WID_KEY -Descending
    [FT_IO]::Write_JSON_Array($exists_json_path, $final_list)

    [PSCustomObject[]]$customized_obj_arr = fn_PSCOjb_Arr $final_list
    [FT_IO]::Write_JSON_Array($customized_wid_path, $customized_obj_arr)
  }
  if (-not(Test-Path $exists_json_path)) {
    [FT_IO]::Write_JSON_Array($exists_json_path, $addition_list)

    [PSCustomObject[]]$customized_obj_arr = fn_PSCOjb_Arr $addition_list
    [FT_IO]::Write_JSON_Array($customized_wid_path, $customized_obj_arr)
  }
}
catch [exception] {
  Write-Host $config.command_name ("エラー😢😢😢 : " + $_)
  Write-Output $_
}
finally {
  $excel.Quit()
  foreach ($_ in @( $Sheet, $book , $excel)) {
    if ($_ -ne $null) {
      [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($_)
    }
  }
  exit 0
}


# コマンド終了
exit 0

