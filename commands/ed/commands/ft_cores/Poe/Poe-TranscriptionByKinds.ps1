Param(
  [Parameter(Mandatory = $true, Position = 0)]
  [hashtable]$_applicants_dict,
  [Parameter(Mandatory = $True, Position = 1)]
  [PSCustomObject]$_config,
  [Parameter(Mandatory = $True, position = 2)]
  [hashtable]$_subject
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"


function fn_Generate_Export_Path {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [ValidatePattern("ts|tp|ih|TS|TP|IH")]
    [String][ref]$_Tgroup,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject][ref]$_config,
    [Parameter(Mandatory = $True, Position = 2)]
    [String][ref]$_applicants_names
  )
  $export_name = @(
    "${HOME}",
    $_config.$_Tgroup.export_folder,
    $_config.File.head_name,
    $_applicants_names,
    $_config.File.extension
  ) -join ""
  return  [FT_Name]::One_Liner($export_name)
}


. .\ft_cores\FT_Name.ps1
$script:kinds = $_subject.kinds
$script:Tgroup = $_subject.Tgroup
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

  foreach ($_applicant in $_applicants_dict.Values) {

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

    $name_field = $_config.File.applicant
    $export_path = fn_Generate_Export_Path ([ref]$_Tgroup) ([ref]$_config) ([ref]$_applicant.$name_field);

    # 空ファイルを作成
    New-Item -Path $export_path -ItemType File -Force
    
    # 空ファイルに書き込む
    Write-Output "🌸🌸🌸  出力先 : $export_path" 
    $book.SaveAs($export_path)
  }
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

