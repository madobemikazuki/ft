Param(
  [Parameter(Mandatory = $true, Position = 0)]
  [System.Collections.Generic.List[PoeObject]]$_posting_object,
  [Parameter(Mandatory = $True, Position = 1)]
  [PSCustomObject]$_poe_config,
  [Parameter(Mandatory = $True, position = 3)]
  [String]$_export_path
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

try {
  # Measure-Command でスクリプトブロック内の実行完了時間を測定できる。
  #$time = Measure-Command {}
  #Write-host $time.TotalSeconds.ToString("F2")"秒 : Excelの起動が完了するまでの経過時間"

  $script:excel = New-Object -ComObject Excel.Application

  #.Visible = $false でExcelを表示しないで処理を実行できる。
  $excel.Visible = $False

  # 上書き保存時に表示されるアラートなどを非表示にする
  $excel.DisplayAlerts = $False

  # リンクの更新方法が 0 の場合は何もしない。
  #.Workbooks.Open(ファイル名, リンクの更新方法, 読み取り専用) でExcelを開きます。
  $script:book = $excel.Workbooks.Open((${Home} + $_poe_config.temp_path), 0, $true)


  <# Worksheets.Item('シート名') で指定したシートを開くときの注意点
      ExcelのエンコードはSJISなので、シート名が日本語のときは、
      PowerShellのファイルはSJISにして実行する必要があります。
      PowerShellのファイルを UTF-8 で保存すると、日本語のシート名が検索できないので、
      代わりに .Worksheets.Item(シート番号) とする方法もあります。
    #>

  $sheet = $book.Worksheets.Item($_poe_config.temp_sheet_page)

  foreach ($_ in $_posting_object) {
    $sheet.Cells.Item($_.point_x, $_.point_y) = $_.value
  }


  # プリントアウトする?
  $print_config = $_poe_config.printing
  if ($print_config.printable) {
    $book.PrintOut.Invoke(
      @(
        [int16]$print_config.start_page,
        [int16]$print_config.end_page,
        [int16]$print_config.number_of_copies
      )
    )
  }    
    
  # 空ファイルを作成し、そこに出力
  New-Item $_export_path -type file -Force
  $book.SaveAs($_export_path)

  Write-Output "👍👍👍  出力先 : $_export_path"    
  $book.Close()
}
catch [exception] {
  Write-Output "😢😢😢エラーをよく読んでね。"
  $error[0].ToString()
  Write-Output $_
}
finally {
  $excel.Quit()
  foreach ($_ in @( $sheet, $book , $excel)) {
    if ($null -ne $_) {
      [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($_)
    }
  }
  exit 0
}

