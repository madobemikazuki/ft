Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [ValidatePattern("ts|tp|ih|TS|TP|IH")]
  [String]$_Tgroup,
  [Parameter(Mandatory = $True, Position = 1)]
  [ValidatePattern("c|d|cd|j|C|D|CD|J")]
  [String]$_kinds,
  [Parameter(Mandatory = $true, Position = 2)]
  [ValidatePattern("^\d{2}\b\-\b\d{6}$")]
  [String[]]$_regist_nums
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"


# TODO: 関数群
function private:fn_Read_Config_JSON {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$_config_path
  )
  Get-Content -Path $_config_path -Encoding utf8 | ConvertFrom-Json
}

function private:fn_Read_CSV {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$_source_path,
    [Parameter(Mandatory = $True, Position = 1)]
    [String]$_encode
  )
  Import-Csv -Path $_source_path -Encoding $_encode
}

function private:fn_SearchTargets {
  Param(
    [Parameter(Mandatory = $True, Position = 0) ]
    [PSCustomObject[]][ref]$_source_array,
    [Parameter(Mandatory = $True, Position = 1) ]
    [ValidatePattern("^\d{2}\b\-\b\d{6}$")]
    [String[]]$_targets,
    [Parameter(Mandatory = $True, Position = 2)]
    [String]$_flag
  )
  # Source から 単一Flag の値に該当するする複数のTargets のオブジェクトリストを返す 
  [PSCustomObject[]]$private:result = $_source_array | Where-Object { $_.$_flag -in $_targets } 
  return $result
}


function private:fn_CombinedObject {
  Param(
    [Parameter(Mandatory = $true, Position = 0)]
    [String[]] $_header,
    [Parameter(Mandatory = $true, Position = 1)]
    [PSCustomObject]$_position,
    [Parameter(Mandatory = $true, Position = 2)]
    [PSCustomObject]$_applicant
  )
  $COLON = ':'
  [Array]$address = foreach ($_ in $_header) {
    $p = [UInt16[]] $_position.$_.split($COLON)
    [PSCustomObject] @{
      name    = $_
      point_x = $p[0]
      point_y = $p[1]
      value   = $_applicant.$_
    }
  }
  return $address
}



# TODO: 初期化処理 --------------------------------------------
[PSCustomObject]$config = fn_Read_Config_JSON ".\config\ed.json"
#$config.TS.export_folder 
Test-Path (${HOME} + $config.gZEN_csv)

[PSCustomObject[]]$source_list = fn_Read_CSV (${HOME} + $config.gZEN_csv) $config.encoding
#$source_list | Format-Table

[PSCustomObject[]]$applicants = fn_SearchTargets ([ref]$source_list) $_regist_nums $config.search_field
#$applicants | Format-Table

$head = $config.address_table.psobject.properties.Name
#$head

[PSCustomObject[]]$application_contents = foreach ($object in $applicants) { $object | Select-Object -Property $head } 
#$application_contents | Format-List


# TODO: 申請用オブジェクト生成 --------------------------------------


# TODO: エクセルオブジェクトへマッピング、書き出し 副作用満載-------------
try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $False
  $excel.DisplayAlerts = $False
  #.Workbooks.Open(ファイル名, リンクの更新方法, 読み取り専用) でExcelを開きます1。
  # リンクの更新方法が 0 の場合は何もしない。
  $book = $excel.Workbooks.Open(
    (${HOME} + $config.template_path),
    0,
    $true
  )

  [int16[]]$pages = $config.$_Tgroup.sheet_pages.$_kinds
  #$pages

  foreach ($content in $application_contents) {
    $formatted_obj = fn_CombinedObject $head $config.address_table $content
    $formatted_obj | Format-table

    foreach ($page in $pages) {
      $sheet = $book.Worksheets.Item($page)
      foreach ($_ in $formatted_obj) {
        $sheet.Cells.Item($_.point_x, $_.point_y) = $_.value
      }
      # プリントアウトする
      $book.PrintOut.Invoke(@($page, $page, [int16]$config.printing.number_of_copies))
    }

    # exportする
    $export_path = @(
      "${HOME}",
      $config.$_Tgroup.export_folder,
      $config.File.head_name,
      $content.($config.File.applicant),
      $config.File.extension
    ) -join ""
    # 空ファイルを作成
    New-Item -Path $export_path -ItemType File -Force
    
    # 空ファイルに書き込む
    $book.SaveAs($export_path)
    #$values | Format-Table   
    #Write-Output "👍👍👍  出力先 : $export_path"    
  }
}
catch [exception] {
  Write-Output "😢😢😢エラーをよく読んでね。"
  $error[0].ToString()
  Write-Output $_
}
finally {
  @($book) | ForEach-Object {
    if ($_ -ne $null) {
      [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($_)
    }
  }
  $excel.Quit()
  [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
}


# コマンド終了
exit 0