Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

# 登録、解除予約情報をJSONファイルに出力するコマンド。
function script:notifycation {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$title,
    [Parameter(Mandatory = $True, Position = 1)]
    [String]$message
  )
  Add-Type -AssemblyName System.Windows.Forms
  $MUTEX_NAME = "Global\mutex" #多重起動チェック用

  try {
    $mutex = New-Object System.Threading.Mutex($False, $MUTEX_NAME)
    #多重起動チェック
    if ($mutex.WaitOne(0, $False)) {
      $notify_icon = New-Object Windows.Forms.NotifyIcon

      $notify_icon.Icon = [Drawing.SystemIcons]::Application
      $notify_icon.Visible = $True

      $notify_icon.BalloonTipText = "$title :  $message"
      $notify_icon.ShowBalloonTip(1)

      # $_second 秒待機して通知を非表示にする。
      $notify_icon.Visible = $False
    }
  }
  finally {
    $notify_icon.Dispose()
    $mutex.ReleaseMutex()
    $mutex.Close()
    $mutex.Dispose()
    exit
  }
}

function script:fn_Read_JSON {
  Param(
    [Parameter(Mandatory = $True)]
    [ValidatePattern('\.json$')]
    [String]$_path
  )
  try {
    $json = Get-Content -Path $_path -Encoding UTF8 | ConvertFrom-Json
    return $json
  }
  catch [exception] {
    $private:message = "エラー😢😢😢 : " + $_path + " が存在しません。" + $_
    Write-Host $message
    notifycation $config.command_name $message
  }
}


function script:fn_Find_Latest_File {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$_Folder,
    [Parameter(Mandatory = $True, Position = 1)]
    [String]$_TargetName
  )
  try {
    # 指定したフォルダ内から、指定した名前を含むファイル群から最新の更新日のファイルパスを返す
    $file_list = Get-ChildItem -Path $_Folder -File -Filter $_TargetName
    $latest_file = ($file_list | Sort-Object LatestWriteTime -Descending)[0].FullName
    return $latest_file
  }
  catch [exception] {
    $private:message = "エラー😢😢😢 : " + $_TargetName + "ファイルが存在しない。" + $_
    Write-Output $message
    notifycation $config.command_name $message
  }
}




function script:fn_YEAR {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$_mmdd,
    [Parameter(Mandatory = $True, Position = 1)]
    [String]$_format
  )
  # 現在時間と引数の日付を比較して過去なら次の年の日付を返す。
  $now = Get-Date
  $this_year = Get-Date $now -Format "yyyy年"
  $date_full = Get-Date ($this_year + $_mmdd)
  # $date_full.gettype()
  switch ($date_full) {
    # スクリプトブロックで比較演算子を使用できる。
    # ()で囲んでいるのは可読性を少しでも担保するためである。
    ({ $now -le $_ }) {
      return $_.ToString($_format)
      #return [datetime]::ParseExact($_, $_format, $null)
      break
    }
    ({ $now -gt $_ }) {
      return $_.AddYears(1).ToString($_format)
      #return [datetime]::ParseExact($_.AddYears(1), $_format, $null)
      break
    }
  }
}

function script:fn_SORT {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject[]]$_list,
    [Parameter(Mandatory = $True, Position = 1)]
    [String[]]$_props
  )
  #スクリプトブロック直接渡せるの便利ですね。
  $new_list = foreach ($_ in $_props) {
    $_list | Sort-Object -Property $_
  }
  return $new_list
}

function script:fn_Append_KV {
  Param(
    [Parameter(mandatory = $True, Position = 0)]
    [PSCustomObject]$_obj,
    [Parameter(Mandatory = $True, Position = 1)]
    [String]$_key,
    [Parameter(Mandatory = $True, Position = 2)]
    [String]$_value
  )
  Add-Member -InputObject $_obj -NotePropertyName $_key -NotePropertyValue $_value -Force
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


[PSCustomObject]$script:config = fn_Read_JSON ".\config\rsv.json"

[PSCustomObject]$target_file_name = $config.common.source_file_name
$source_folder = (${Home} + $config.common.source_folder)
$script:source_file = fn_Find_Latest_File $source_folder $target_file_name
#$file

try {
  $script:excel = New-Object -ComObject Excel.Application
  $excel.Visible = $False
  $excel.DisplayAlerts = $False

  $script:book = $excel.Workbooks.Open($source_file, 0, $True)

  # TODO: ここから二通りの処理をしたい
  # 1 は登録予定の処理
  # 2 は解除予定の処理
  foreach ($_task in [PScustomObject[]]$config.tasks) {
    $_task.task_name
    $target_page = $_task.sheet_page
    $script:Sheet = $book.Sheets.Item($target_page)

    #行 ( rows y軸) の設定 値取得対象の始点行から最終行までのRange設定
    $starting_row = $_task.starting_values_row
    $end_of_rows = $Sheet.UsedRange.Rows.Count
    Write-Host ("最終行: " + $end_of_rows)
    $select_rows_range = @($starting_row..$end_of_rows)
    Write-Host ("想定している予約の数: " + $select_rows_range.Length)

    #$Sheet のカラムは 1 から始まる。
    $column_num_list = @($_task.starting_column..$_task.end_column)

    Write-host ($_task.task_name + " の フィールド数: " + $column_num_list.Length)

    #フィールド列の値を配列に格納
    $field_row = $_task.field_row
    $field_culumns = foreach ($_column_number in $column_num_list) {
      $Sheet.Cells.Item($field_row, $_column_number).Text
    }

    # JavaScript でも日付情報を利用できるように実用的なフォーマットにはしない。
    # 実用的な "yyyy年　　MM月　　dd日"のような空白文字併記のフォーマットは利用者が用意すべきである。
    $date_format = $config.common.date_format
    $reserved_date_field = $field_culumns -match $config.common.flag_reserved_date
    $reserved_time_field = $field_culumns -match $config.common.flag_reserved_time

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
          $config.common.flag_reserved_date { 
            $_date = $Sheet.Cells.Item($_row, $_column_number).Text
            if ([String]::IsNullOrEmpty($_date)) { break }
            fn_YEAR $_date $date_format
            continue
          }
          Default { $Sheet.Cells.Item($_row, $_column_number).Text }
        }
        Add-Member -InputObject $row_obj -NotePropertyName $_key -NotePropertyValue $_value -Force
      }
      # 予約日が空の $obj は返さない。
      $is_Null_or_Empty = [String]::IsNullOrEmpty($row_obj.$reserved_date_field)
      if ($is_Null_or_Empty) { continue }
      #if (([String]::IsNullOrEmpty($obj.$reserved_date_field))) { continue }

      # 解除申請書の 申請書作成欄 が "済" の $obj は返さない。
      $isCreated = ($target_page -eq 2) -and ($row_obj."申請書作成" -eq $_task.exclusion_value)
      $isIHI = ($target_page -eq 2) -and ($row_obj."管理会社".Contains("IHI元請"))
      if ($isCreated -or $isIHI) { continue }
      $row_obj
    }
  

    # 配列のオブジェクト を 予約時間、予約日の順で昇順ソートする。
    # -Property のkv を複数指定できるよ。HushTableでもok。
    $sorted_reserved_list = $reserved_list | Sort-Object -Property $reserved_time_field | Sort-Object -Property $reserved_date_field
    write-host ("予約済みの数: " + $sorted_reserved_list.Length)
    #$sorted_reserved_list = fn_SORT $reserved_list @($reserved_time_field, $reserved_date_field)
    # JSON出力 JSON出力はUTF8-bomでOK
    # JSONファイルをブラウザ上で読み込んだ場合、
    # 特定の文字列を検索するには F3 が有効である。
    $utf8_with_BOM = New-Object System.Text.UTF8Encoding $True
    $json_path = (${HOME} + $_task.export_path)
    fn_Write_JSON $json_path $sorted_reserved_list $utf8_with_BOM
  }
  #$source_file

  # 参照した事前承認ファイルを \waste へ移動する。
  #Move-Item -Path $source_file -Destination ($source_folder + $config.common.waste_folder)
  notifycation $config.command_name ("出力完了💩💩💩      by ver" + "0.1")

}
catch [exception] {
  notifycation $config.command_name ("エラー😢😢😢 : " + $_)
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

