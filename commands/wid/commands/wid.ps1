﻿Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

<#
  オブジェクト形式で保存  \TEMP\WID_LookUpHash.json
  配列形式で保存  \PAN\WID_min_UTF8-bom.json

  高速化するには配列の使用を止めるべきだが、
  リファクタリングする暇がない。
#>



#. .\core\json\read_json.ps1
function script:fn_Read_JSON {
  Param(
    [Parameter(Mandatory = $True)]
    [ValidatePattern('\.json$')]$_path
  )
  $json = Get-Content -Path $_path -Encoding UTF8 | ConvertFrom-Json
  return $json
}

function script:fn_Find_Latest_File {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$_Folder,
    [Parameter(Mandatory = $True, Position = 1)]
    [String]$_TargetName
  )
  # 指定したフォルダ内から、指定した名前を含むファイル群から最新の更新日のファイルパスを返す
  $file_list = Get-ChildItem -Path $_Folder -File -Filter $_TargetName
  $latest_file = ($file_list | Sort-Object LatestWriteTime -Descending)[0].FullName
  return $latest_file
}

function script:fn_Write_Array_JSON {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$_path,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject[]]$_Object_List,
    [Parameter(Mandatory = $True, Position = 2)]
    [System.Text.Encoding]$_encoding
  )
  #既存するファイルを上書きする
  if (Test-Path $_path) {
    New-Item -Path $_path -ItemType File -Force
  }
  [System.IO.File]::WriteAllLines($_path, (ConvertTo-Json $_Object_List), $_encoding)
}

function script:fn_Write_JSON {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$_path,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject]$_Object,
    [Parameter(Mandatory = $True, Position = 2)]
    [System.Text.Encoding]$_encoding
  )
  #既存するファイルを上書きする
  if (Test-Path $_path) {
    New-Item -Path $_path -ItemType File -Force
  }
  [System.IO.File]::WriteAllLines($_path, (ConvertTo-Json $_Object), $_encoding)
}

function script:fn_notifycation {
  Param(
    [String]$title,
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

function fn_LookUpHash {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject[]]$_Object_List,
    [Parameter(Mandatory = $True, Position = 1)]
    [String]$_key
  )
  $LookUpHash = [ordered]@{}
  . .\utils\ToWide.ps1
  foreach ($_ in $_Object_List) {
    $new_obj = @{}
    $new_obj["subject"] = $_."作業件名"
    # -replace 演算子は正規表現を利用できる
    $private:group_name = $_."作業主管グループ" -replace '[ＧG]$', 'グループ'
    $new_obj["depertment"], $new_obj["group"] = (ToWide $group_name).Split("　")
    if ($null -eq $new_obj["group"]) { $new_obj["group"] = "" }
    $new_obj["wid"] = $_.$_key
    $LookUpHash[$_.$_key] = $new_obj
  }
  return [PSCustomObject]$LookUpHash
}


$config = fn_Read_JSON ".\config\wid_group.json"

# 対象フォルダー
$source_folder = (${HOME} + $config.import.folder)
$name = $config.import.contained_name
$script:wid_XLS_file_path = fn_Find_Latest_File $source_folder $name

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
  [PScustomObject[]]$list = foreach ($_row in $select_rows_range) {
    # pscustomObject に格納する。
    # return object
    $private:obj = @{}
    foreach ($_column in $columns) {
      $index = $columns.IndexOf($_column)
      $key = $export_field[$index]
      $value = $Sheet.Cells.Item($_row, $_column).Text
      $obj[$key] = $value
    }
    [PSCustomObject]$obj
  }
  $script:WID_KEY = $config.export.wid_key
  $new_list = $list | Sort-Object -Property $WID_KEY -Descending
  $regexp = $config.export.wid_regexp
  $final_list = $new_list | Where-Object { $_.$WID_KEY -match $regexp }

  #$final_list.length
  #$final_list | Format-Table

  # JSON出力 JSON出力はUTF8-bomでOK
  # JSONファイルをブラウザ上で読み込んだ場合、
  # 特定の文字列を検索するには F3 が有効である。
  $utf8_with_BOM = New-Object System.Text.UTF8Encoding $True

  # 最小限の情報を JSON に出力する
  $min_selcets = $config.export.min_field
  [PSCustomObject[]]$script:addition_list = foreach ($_ in $final_list) {
    $_ | Select-Object -Property $min_selcets
  }
  #Write-Host "addition_list 項目追加されている新しいオブジェクトのリスト"
  #Write-Host "xlsから読み込んだもの"
  #$addition_list | Format-Table
  $exists_json_path = (${HOME} + $config.export.min_json_path)
  $LookUpHash_path = (${HOME} + $config.export.LookUpHash_path)

  # $addition_list が完成したのち、既存の WID_min_UTF8-bom.json
  if (Test-Path $exists_json_path) {
    Write-Host "Values"
    $exists_obj_list = fn_Read_JSON $exists_json_path
    $exists_obj_list | Format-Table
    Write-Host "問題なし"
    . .\ft_core\ADF.ps1
    [PSCustomObject[]]$script:complete_list = ADF $exists_obj_list $addition_list $WID_KEY
    if (($complete_list[0] -eq 0) -or ($Null -eq $complete_list)) {
      fn_Notifycation $config.command_name "追記するものなし。更新せずに終了しました。"
      exit 0
    }

    $final_list = $complete_list | Sort-Object -Property $WID_KEY -Descending
    fn_Write_Array_JSON $exists_json_path $final_list $utf8_with_BOM

    $LookUpObject = fn_LookUpHash $final_list $WID_KEY
    fn_Write_JSON $LookUpHash_path $LookUpObject $utf8_with_BOM

    fn_Notifycation $config.command_name ("出力完了💩💩💩      Ver " + $config.version)

  }
  if (-not(Test-Path $exists_json_path)) {
    fn_Write_Array_JSON $exists_json_path $addition_list $utf8_with_BOM

    $LookUpObject = fn_LookUpHash $addition_list $WID_KEY
    fn_Write_JSON $LookUpHash_path $LookUpObject $utf8_with_BOM

    fn_Notifycation $config.command_name ("新規出力完了💩💩💩      Ver " + $config.version)
  }
}
catch [exception] {
  fn_Notifycation $config.command_name ("エラー😢😢😢 : " + $_)
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

