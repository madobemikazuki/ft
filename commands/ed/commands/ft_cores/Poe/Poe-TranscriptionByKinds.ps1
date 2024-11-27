Param(
  [Parameter(Mandatory = $true, Position = 0)]
  [PSCustomObject]$_applicant,
  [Parameter(Mandatory = $True, Position = 1)]
  [PSCustomObject]$_config,
  [Parameter(Mandatory = $True, position = 2)]
  [hashtable]$_subject,
  [Parameter(Mandatory = $True, position = 3)]
  [String]$_export_path
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"


$script:kinds = $_subject.kinds
$script:Tgroup = $_subject.Tgroup
$script:export_path = $_export_path
$script:template_path = $_subject.template

try {

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $False
  $excel.DisplayAlerts = $False
  #.Workbooks.Open(ファイル名, リンクの更新方法, 読み取り専用) でExcelを開きます1。
  # リンクの更新方法が 0 の場合は何もしない。
  $book = $excel.Workbooks.Open(
    #(${HOME} + $config.template_path),
    $template_path,
    0,
    $true
  )


  foreach ($_kind in $kinds.ToCharArray()) {
    $page = $_config.$Tgroup.sheet_pages.$_kind
    $formatted_obj = [PoeAddress]::Single_Format($_applicant, $_config.$_kind.address_table)
    $formatted_obj[0].value = ($_config.$_kind.sandwitch) -replace @($_config.replacement, $formatted_obj[0].value)
    #$formatted_obj | Format-Table
      
    $sheet = $book.Worksheets.Item($page)
    foreach ($_obj in $formatted_obj) {
      $sheet.Cells.Item($_obj.point_x, $_obj.point_y) = $_obj.value
    }
    #プリントアウトする
    if ($_config.printable) {
      $book.PrintOut.Invoke(@($page, $page, [int16]$_config.printing.number_of_copies))
    }
  }



  # 空ファイルを作成
  New-Item -Path $export_path -ItemType File -Force
    
  # 空ファイルに書き込む
  Write-Output "🌸🌸🌸  出力先 : $export_path" 
  $book.SaveAs($export_path)

  $excel.Quit()
}
catch [exception] {
  Write-Output "😢😢😢エラーをよく読んでね。"
  $error[0].ToString()
  Write-Output $_
  $excel.Quit()
}
finally {
  $excel.Quit()
  $excel = $null
  [System.GC]::Collect()
  foreach ($_ in @( $sheet, $book , $excel)) {
    if ($null -ne $_) {
      [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($_)
    }
  }
  [System.GC]::Collect()
}

