# 登録、解除予約情報をJSONファイルに出力するコマンド。

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

# 使用する自作コマンドレット
. ..\ft_cores\FT_IO.ps1
. ..\ft_cores\FT_Date.ps1
. ..\ft_cores\FT_Dict.ps1
. ..\ft_cores\FT_Array.ps1

$config_path = ".\config\FT_Utils.json"
$command_name = Split-Path -Leaf $PSCommandPath

[PSCustomObject]$script:config = ([FT_IO]::Read_JSON_Object($config_path)).$command_name

[PSCustomObject]$target_file_name = $config.common.source_file_name
$source_folder = (${Home} + $config.common.source_folder)
$script:source_file = [FT_IO]::Find_Latest_File($source_folder, $target_file_name)

try {
  $script:excel = New-Object -ComObject Excel.Application
  $excel.Visible = $False
  $excel.DisplayAlerts = $False

  $script:book = $excel.Workbooks.Open($source_file, 0, $True)


  $LookUp_task_dict = @{}
  # 1 は登録予定の処理
  # 2 は解除予定の処理
  foreach ($_task in [PScustomObject[]]$config.tasks) {
    Write-Host "[" $_task.task_name "]"
    $target_page = $_task.sheet_page
    $script:Sheet = $book.Sheets.Item($target_page)


    # TODO: 読込開始地点の定義
    # 今日の日付、もしくは今日の日付以降の行を取得それ以外は無視する。
    # 実行速度がかなり遅い？
    #行 ( rows y軸) の設定 値取得対象の始点行から最終行までのRange設定
    $starting_row = $_task.starting_values_row
    $end_of_rows = $Sheet.UsedRange.Rows.Count
    Write-Host ("最終行: " + $end_of_rows)
    $select_rows_range = @($starting_row..$end_of_rows)
    #Write-Host ("想定している予約の数: " + $select_rows_range.Length)

    #$Sheet のカラムは 1 から始まる。
    $column_num_list = @($_task.starting_column..$_task.end_column)

    #Write-host ($_task.task_name + " の フィールド数: " + $column_num_list.Length)

    #フィールド列の値を配列に格納
    $field_row = $_task.field_row
    $field_culumns = foreach ($_column_number in $column_num_list) {
      $Sheet.Cells.Item($field_row, $_column_number).Text
    }
    


    # JavaScript でも日付情報を利用できるように実用的なフォーマットにはしない。
    # 実用的な "yyyy年　　MM月　　dd日"のような空白文字併記のExcel的なフォーマットは利用者が用意すべきである。
    $date_format = $config.common.date_format
    $reserved_search_key = $field_culumns -match $config.common.search_key
    $reserved_date_field = $field_culumns -match $config.common.reserved_date
    $reserved_time_field = $field_culumns -match $config.common.reserved_time
    $today = Get-Date

    [PScustomObject[]]$reserved_list = foreach ($_row in $select_rows_range) {
      # $Sheetの行ごとのカラムを pscustomObject に格納する。
      $private:row_obj = [PSCustomObject]@{}
      foreach ($_column_number in $column_num_list) {
        
        $_index = $column_num_list.IndexOf($_column_number)
        $_key = $field_culumns[$_index]
        
        # $_key に予約日が含まれるなら 西暦付きの日付に修正する
        # switch に -Regex オプションを付記すると case部に指定した文字列に一致する右辺を実行する
        $_value = switch -Regex ($_key) {
          # 単語だけ指定すると含まれるものとして評価される  
          # 予約日フィールドに値があれば今日以降の日付であれば値を取得する。
          $config.common.reserved_date { 
            $_value_date = $Sheet.Cells.Item($_row, $_column_number).Text
            if ([String]::IsNullOrEmpty($_value_date)) { 
              [FT_Date]::Ja_Empty_Format()
              break
            }
            #fn_TDTW $today $_value_date $date_format
            [FT_Date]::From_Today_Onwards($today, $_value_date, $date_format)
            break
          }
          Default { $Sheet.Cells.Item($_row, $_column_number).Text }
        }
        Add-Member -InputObject $row_obj -NotePropertyMembers @{ $_key = $_value } -Force
      }

      # 多数の人員が登録しにくると予約日が空でも最低限の情報が必要になる
      $is_Null_or_Empty_list = @(
        [String]::IsNullOrEmpty($row_obj.$reserved_search_key)
        #[String]::IsNullOrEmpty($row_obj.$reserved_date_field),
        #[String]::IsNullOrEmpty($row_obj.$reserved_time_field)
      )
      if ($is_Null_or_Empty_list -contains $True) { continue }

      # 解除申請書の 申請書作成欄 が "済" の $obj は返さない。
      # $isCreated = ($target_page -eq 2) -and ($row_obj."申請書作成" -eq $_task.exclusion_value)
      $isIHI = ($target_page -eq 2) -and ($row_obj."管理会社".Contains("IHI元請"))
      #if ($isCreated -or $isIHI) { continue }
      if ($isIHI) { continue }
      $row_obj
    }  

    # 配列のオブジェクト を 予約時間、予約日の順で昇順ソートする。
    # -Property のkv を複数指定できるよ。HushTableでもok。
    $sort_keys = @($reserved_time_field, $reserved_date_field)
    $sorted_reserved_list = [FT_Array]::Sort($reserved_list, $sort_keys)
    write-host ("予約済みの数: " + $sorted_reserved_list.Length)
    # JSON出力 JSON出力はUTF8-bomでOK
    # JSONファイルをブラウザ上で読み込んだ場合、特定の文字列を検索するには F3 が有効である。
    
    $LookUpHash = [FT_Array]::ToDict($sorted_reserved_list, $reserved_search_key)
    $selected_LookUpHash = [FT_Dict]::Selective($LookUpHash, $_task.selection)
    $LookUp_task_dict[$_task.task_initial] = $selected_LookUpHash
    #$json_path = (${HOME} + $config.common.export_folder + $_task.export_name)
    #[FT_IO]::Write_JSON_Object($json_path, [PSCustomObject]$selected_LookUpHash)
  }
  $json_path = (${HOME} + $config.common.export_folder + $config.common.export_file)
  [FT_IO]::Write_JSON_Object($json_path, [PSCustomObject]$LookUp_task_dict)
  Write-Host ($config.command_name + "::出力完了💩💩💩   by Ver " + $config.version)
}
catch [exception] {
  Write-Host ($config.command_name + "::エラー😢😢😢")
  Write-Output $_
}
finally {
  $excel.Quit()
  foreach ($_ in @( $Sheet, $book , $excel)) {
    if ($null -ne $_) {
      [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($_)
    }
  }
}

