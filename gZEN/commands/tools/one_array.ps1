<#
.SYNOPSIS
  Excel ファイルから一行分の値を印刷し、JSONにも出力するコマンド
.DESCRIPTION
  Excel ファイルを他のプログラムで再利用するときに必要なフィールド名を得るためのもの。
  CSV や JSON のフィールド名に援用できる。
.EXAMPLE
  .\one_array.ps1 *.xls 1
.EXAMPLE
  このコマンドレットの使用方法の別の例
.INPUTS
  第一引数にExcelファイルパスを指定する。
  第二引数に値を取り出したい行を指定する。
.OUTPUTS
  出力結果をプリンタから出力する
  出力結果をJSONに出力する。
.NOTES
  全般的な注意
.COMPONENT
  このコマンドレットが属するコンポーネント
.ROLE
  このコマンドレットが属する役割
.FUNCTIONALLITY
  このコマンドレットの機能
#>
Param(
  [Parameter(Mandatory = $True, position = 0)]
  [String]$_excel_path,
  [Parameter(Mandatory = $True, Position = 1)]
  [int16]$_row
)



Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

function script:fn_Write_JSON {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$_path,
    [Parameter(Mandatory = $True, Position = 1)]
    [Object[]]$_Object_List,
    [Parameter(Mandatory = $True, Position = 2)]
    [System.Text.Encoding]$_encoding
  )
  if (Test-Path $_path) { New-Item -Path $_path -ItemType File -Force }
  [System.IO.File]::WriteAllLines($_path, (ConvertTo-Json $_Object_List), $_encoding)
}

if (!(Test-Path $_excel_path)) {
  Write-Host "この世に悪は存在しません。あなたの提示したファイルがこの具象世界には存在しないように。"
  exit 0
}

try {
  $script:excel = New-Object -ComObject Excel.Application
  #.Visible = $false でExcelを表示しないで処理を実行できる。
  $excel.Visible = $False
  # 上書き保存時に表示されるアラートなどを非表示にする
  $excel.DisplayAlerts = $False

  # リンクの更新方法が 0 の場合は何もしない。
  #.Workbooks.Open(ファイル名, リンクの更新方法, 読み取り専用) でExcelを開くことができる。
  $script:book = $excel.Workbooks.Open($_excel_path, 0, $True)
  $sheet = $book.Worksheets.Item(1)
  
  $start_column = 1
  # SpecialCells() に11を渡すと末尾の数が返ってくる
  [int16]$end_of_column = $sheet.Rows.Item($_row).Cells.SpecialCells(11).Column
  #Write-Host $end_of_column
  $x_list = @($start_column..$end_of_column)


  [String[]]$row_columns = foreach ($_ in $x_list) {
    $sheet.Rows.Item($_row).Cells.Item($start_column, $_).Value2
  }
  #$row_columns.GetType()
  
  # そのまま印刷する。
  $row_columns | Out-Printer

  $utf8_with_BOM = New-Object System.Text.UTF8Encoding $True
  $export_path = ("${HOME}" + "\Downloads\PAN\フィールド.json")
  fn_Write_JSON $export_path $row_columns $utf8_with_BOM
}
catch [exception] {
  Write-Output "😢😢😢エラーをよく読んでね。"
  $error[0].ToString()
  Write-Output $_
  exit 404
}
finally {
  $excel.Quit()
  foreach ($_ in @($sheet, $book, $excel)) {
    if ($null -ne $_) {
      [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($_)
    }
  }
  exit 0
}

