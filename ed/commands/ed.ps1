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




# 今から作っていくところ   --------------
<#
 ローカルスコープ上に作っていく。
#>

# TODO: 関数群
function private:fn_Read_JSON {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$_config_path
  )
  Get-Content -Path $_config_path | ConvertFrom-Json
}

function private:fn_Read_CSV {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$_source_path
  )
  Import-Csv -Path $_source_path -Encoding UTF8
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

function private:fn_CombinedName {
  Param(
    [Parameter(Mandatory = $true, Position = 0)][String]$first_name,
    [Parameter(Mandatory = $true, Position = 1)][String]$last_name,
    [String]$delimiter = '　'#デフォルト引数 必要な時は呼び出し側で -delimiter を指定すること
  )

  $sb = New-Object System.Text.StringBuilder
  #副作用処理  StringBuilderならちょっと速いらしい。要素数が少ないから意味ないかも。
  @($first_name, $delimiter , $last_name) | ForEach-Object { [void] $sb.Append($_) }
  return $sb.ToString()
}

function private:fn_YoureCompanyNames {
  Param(
    [Parameter(Mandatory = $True, Position = 0)][String]$_managemanet_com_name,
    [Parameter(Mandatory = $True, Position = 1)][String]$_employer_name
  )

  if ($_managemanet_com_name -eq $_employer_name) {
    return $_managemanet_com_name
  }
  # 二つの名前が違うとき実行
  if (!($_managemanet_com_name -eq $_employer_name)) {
    return fn_CombinedName $_managemanet_com_name $_employer_name  -delimiter " / "
  }
}

function private:fn_ShotenComType {
  Param(
    [Parameter(Mandatory = $True)]
    [String]$_corporate_name
  )
  switch ($_corporate_name) {
    { $_.Contains('株式会社') } { return $_.Replace('株式会社', '（株）') }
    { $_.Contains('有限会社') } { return $_.Replace('有限会社', '（有）') }
  }
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
[PSCustomObject]$config = fn_Read_JSON ".\config\ed.json"
#$config.TS.export_folder 

[PSCustomObject[]]$source_list = fn_Read_CSV (${HOME} + $config.gZEN_csv)
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

    # exportする。
    $export_path = @(
      "${HOME}",
      $config.$_Tgroup.export_folder,
      $config.File.head_name,
      $content.($config.File.applicant),
      $config.File.extension
    ) -join ""
    New-Item -Path $export_path -ItemType File -Force

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
