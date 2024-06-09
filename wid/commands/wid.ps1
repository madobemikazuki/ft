Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

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

function script:notifycation {
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

$config = fn_Read_JSON ".\config\wid_group.json"

# 対象フォルダー
$source_folder = (${HOME} + $config.import.folder)
$name = $config.import.contained_name
$script:wid_file = fn_Find_Latest_File $source_folder $name
$wid_file

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $False
  $excel.DisplayAlerts = $False

  $book = $excel.Workbooks.Open($wid_file, 0, $True)
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
    $obj = [PSCustomObject]@{}
    foreach ($_column in $columns) {
      $index = $columns.IndexOf($_column)
      $_key = $export_field[$index]
      $value = $Sheet.Cells.Item($_row, $_column).Text
      $obj | Add-Member -MemberType NoteProperty -Name $_key -Value $value
    }
    $obj
  }

  $new_list = $list | Sort-Object -Property '作業件名コード' -Descending

  $key = $config.export.wid_key
  $regexp = $config.export.wid_regexp
  $final_list = $new_list | Where-Object { $_.$key -match $regexp }  
  $final_list.length
  #$final_list | Format-Table


  # csv出力
  # csvに出力するには、エンコードはANSIでなければならない。
  # csvなんて使わないから コメントアウトしておく
  <#
  $csv_path = "${HOME}" + $config.export.csv_path
  New-Item -Path $csv_path -ItemType File -Force
  $new_list | Export-csv -path $csv_path -NoTypeInformation -Encoding Default
  #>

  # JSON出力 JSON出力はUTF8-bomでOK
  # JSONファイルをブラウザ上で読み込んだ場合、
  # 特定の文字列を検索するには F3 が有効である。
  $utf8_with_BOM = New-Object System.Text.UTF8Encoding $True
  
  # fn_Write_JSON に指定する引数の型を調べるのに使用
  #$utf8_with_BOM.gettype()

  <#
  $json_path = "${HOME}" + $config.export.json_path
  fn_Write_JSON $json_path $final_list $utf8_with_BOM
  #>


  # 最小限の情報を JSON に出力する
  $selcets = $config.export.min_field
  $min_list = foreach ($_ in $final_list) {
    $_ | Select-Object -Property $selcets
  }
  $min_json_path = ("${HOME}" + $config.export.min_json_path)
  fn_Write_JSON $min_json_path $min_list $utf8_with_BOM


  notifycation $config.command_name ("出力完了💩💩💩      Ver " + $config.version)
}
catch [exception] {
  notifycation $config.command_name ("エラー😢😢😢 : " + $_)
  Write-Output $_
}
finally {
  $excel.Quit()
  foreach ($_ in @( $Sheet, $book , $excel)) {
    if ($_ -ne $null) {
      [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($_)
    }
  }
}

# コマンド終了
exit 0

