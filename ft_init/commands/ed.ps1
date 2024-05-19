Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [ValidatePattern("ts|tp|ih|TS|TP|IH")]
  [String]$_Tgroup,
  [Parameter(Mandatory = $True, Position = 1)]
  [ValidatePattern("c|d|cd|j|C|D|CD|J")]
  [String]$_Kinds,
  [Parameter(Mandatory = $true, Position = 2)]
  [ValidatePattern("^\d{2}\b\-\b\d{6}$")]
  [String[]]$_regist_nums
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"



# TODO: 関数群
function private:fn_Read_JSON {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$_path,
    [Parameter(Mandatory = $True, Position = 1)]
    [String]$_encode
  )
  Get-Content -Path $_path -Encoding $_encode | ConvertFrom-Json
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


function script:fn_Shorten_Com_Type_Name {
  Param(
    [Parameter(Mandatory = $True)]
    [String]$_corporate_name
  )
  switch ($_corporate_name) {
    { $_.Contains('株式会社') } { return $_.Replace('株式会社', '（株）') }
    { $_.Contains('有限会社') } { return $_.Replace('有限会社', '（有）') }
    default { return $_corporate_name }
  }
}

function script:fn_Combined_Name {
  Param(
    [Parameter(Mandatory = $true, Position = 0)][String]$_first_name,
    [Parameter(Mandatory = $true, Position = 1)][String]$_last_name,
    [String]$_delimiter = '　'
  )

  $sb = New-Object System.Text.StringBuilder
  #副作用処理  StringBuilderならちょっと速いらしい。要素数が少ないから意味ないかも。
  @($_first_name, $_delimiter , $_last_name) | ForEach-Object { [void] $sb.Append($_) }
  return $sb.ToString()
}

function script:fn_Application_Company_Names {
  Param(
    [Parameter(Mandatory = $True, Position = 0)][String]$_managemanet_com_name,
    [Parameter(Mandatory = $True, Position = 1)][String]$_employer_name
  )
  if ($_managemanet_com_name -eq $_employer_name) {
    return $_managemanet_com_name
  }
  # 二つの名前が違うとき実行
  if (!($_managemanet_com_name -eq $_employer_name)) {
    return fn_Combined_Name $_managemanet_com_name $_employer_name  -_delimiter " / "
  }
}

function private:fn_Extract_Registered_Targets {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject[]][ref]$_registered_persons
  )
  [PSCustomObject[]]$list = foreach ($_registered in $_registered_persons) {
    $shorten_1 = fn_Shorten_Com_Type_Name $_registered."電力申請会社名称"
    $shorten_2 = fn_Shorten_Com_Type_Name $_registered."雇用名称"
    $full_name = fn_Combined_Name $_registered."氏名（姓）" $_registered."氏名（名）"
  
    [PSCustomObject]@{
      "登録時申請会社" = fn_Application_Company_Names $shorten_1 $shorten_2
      "カタカナ氏名"  = $_registered."氏名（カナ）"
      "漢字氏名"    = $full_name 
      "中央登録番号"  = $_registered."中登番号"
    }
  }
  return $list
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


# 初期化処理 --------------------------------------------
[PSCustomObject]$config = fn_Read_JSON ".\config\ed.json" "UTF8"
#$config.TS.export_folder 
#Test-Path (${HOME} + $config.gZEN_csv)

$gZEN_Path = (${HOME} + $config.gZEN_csv)

[PSCustomObject[]]$source_list = fn_Read_CSV $gZEN_path $config.csv_encoding

# 申請用オブジェクト生成 --------------------------------------
# TODO: fn_SearchTargets 関数内で 中央登録番号と合致する人物を探す
# この処理をswitchで探索するか。
[PSCustomObject[]]$applicants = fn_SearchTargets ([ref]$source_list) $_regist_nums $config.search_field
#$applicants | Format-Table


$registered_persons_path = (${home} + $config.registered_list_csv)
[PSCustomObject[]]$registered_list = fn_Read_CSV $registered_persons_path default
[PSCustomObject[]]$script:registered_persons = fn_SearchTargets ([ref]$registered_list) $_regist_nums "中登番号"
if ($null -ne $registered_persons) {
  $applicants = $applicants + @(fn_Extract_Registered_Targets ([ref]$registered_persons))
}



$kinds = $_Kinds.ToCharArray()

# $kinds[0] は ()で囲まなければアクセスできない。
$header = $config.($kinds[0]).address_table.psobject.properties.Name

[PSCustomObject[]]$applications = foreach ($_ in $applicants) { $_ | Select-Object -Property $header } 
#$applications | Format-List



# エクセルオブジェクトへマッピング、書き出し 副作用満載-------------
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

  foreach ($_applicant in $applications) {

    foreach ($_ in $kinds) {
      $page = $config.$_Tgroup.sheet_pages.$_

      $formatted_obj = fn_CombinedObject $header $config.$_.address_table $_applicant
      $formatted_obj[3].value = $config.$_.sandwitch -replace $config.replacement, $formatted_obj[3].value
      $formatted_obj | Format-Table
      
      $sheet = $book.Worksheets.Item($page)
      foreach ($_obj in $formatted_obj) {
        $sheet.Cells.Item($_obj.point_x, $_obj.point_y) = $_obj.value
      }
      #プリントアウトする
      #$book.PrintOut.Invoke(@($page, $page, [int16]$config.printing.number_of_copies))
    }

    # exportする
    $export_path = @(
      "${HOME}",
      $config.$_Tgroup.export_folder,
      $config.File.head_name,
      $_applicant.($config.File.applicant),
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
  #$error[0].ToString()
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

