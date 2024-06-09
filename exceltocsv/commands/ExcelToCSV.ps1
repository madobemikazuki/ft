Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"


function private:fn_Remove_File {
  Param(
    [Parameter(Mandatory = $True, position = 0)]
    [String]$_path,
    [Parameter(Mandatory = $True, position = 1)]
    [String]$_folder
  )
  Move-Item -Path $_path -Destination $_folder -Force
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
      #$ApplicationContext = New-Object System.Windows.Forms.ApplicationContext

      $notify_icon.Icon = [Drawing.SystemIcons]::Application
      $notify_icon.Visible = $True
      # 通知用 のアイコン情報
      #$notify_icon.BalloonTipIcon = [Windows.Forms.ToolTipIcon]::Info

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

function private:fn_Find_Excel_files {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$_Folder
  )
  $private:excel_filename_list = Get-ChildItem -Path $_Folder -File -Name -Include "*.xls", "*.xlsx"
  if ($null -eq $excel_filename_list) {
    fn_notifycation "エラー : Excel ファイルが存在しません"
    exit 404
  }
  return $excel_filename_list
}
function private:fn_Read {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [ValidatePattern("\.csv$|\.json$")]$_path
  )
  switch -Regex ($_path) {
    "\.csv$" {
      return Import-Csv -Path $_path -Encoding Default
    }
    "\.json$" {
      return Get-Content -Path $_path -Encoding UTF8 | ConvertFrom-Json
    }
    Default {
      Write-Host "拡張子が該当しないので終了。"
      exit 0
    }
  }
}


$script:config = fn_Read ".\config\ExcelToCSV.json"
$script:source_fld = (${HOME} + $config.com_fld)
$target_page = 1

[String[]]$filename_list = fn_Find_Excel_files $source_fld

$script:excel = New-Object -ComObject Excel.Application
$excel.Visible = $False
$excel.DisplayAlerts = $False
try {
  foreach ($_filename in $filename_list) {
    $excel_path = $source_fld + $_filename
    $excel_path
    $Book = $excel.Workbooks.Open($excel_path, 0, $True)
    $Sheet = $Book.Sheets.Item($target_page)

    $export_file_name = $_filename -replace "\.xls$|\.xlsx$", ".csv"
    $export_csv_path = (${HOME} + $config.export_folder + $export_file_name)
    # Excel シートのセルに日本語が含まれていると自動的にエンコードが ANSI で出力される。
    $Book.SaveAs($export_csv_path, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlCSV)
    fn_Remove_File $excel_path ($source_fld + "waste")
  }
  fn_notifycation "toCSV" ("出力完了")
}
catch [exception] {
  fn_notifycation "Excel to CSV" ("例外発生😢😢😢 : " + $_)
  Write-Host $_
}
finally {
  $excel.Quit()
  foreach ($_ in @( $Sheet, $Book , $excel)) {
    if ($_ -ne $null) {
      [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($_)
    }
  }
}

$excel.Quit()
exit 0


#$export_json_path = $export_csv_path -replace "\.csv$", ".json"

#[PSCustomObject[]]$csv = Get-Content -Path $export_csv_path -Encoding Default | ConvertFrom-Csv
#$csv
<#
  $end_of_rows = $Sheet.Cells.SpecialCells(11).Row
  # uint16 (0..65535)
  [uint16[]]$row_range = @(($_obj.field_row + 1)..$end_of_rows)

  Write-Host "最終行 : $row_range"

  $end_column = $Sheet.Cells.SpecialCells(11).Column
  #Write-Host "機能してるか？ : $end_column"

  #byte (0..255)
  [byte[]]$column_range = @($_obj.start_column_number..$end_column)
  Write-Host "カラムのレンジ : $column_range"

  $field_columns = foreach ($_ in $column_range) {
    $Sheet.Cells.Item($_obj.field_row, $_).Text
  }
  $field_columns
  #>

