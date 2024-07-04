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


# 関数群

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
    { $_ -match "株式会社|\(株\)" } { return $_ -replace "株式会社|\(株\)", "（株）" }
    { $_ -match "有限会社|\(有\)" } { return $_ -replace "有限会社|\(有\)", "（有）" }
    default { return $_ }
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
    [Parameter(mandatory = $True, Position = 0)]
    [PSCustomObject]$_applicant
  )
  $employment_com = fn_Shorten_Com_Type_Name $_applicant.'雇用名称'
  $manegement_com = fn_Shorten_Com_Type_Name $_applicant.'管理会社名称'
  $application_com = fn_Shorten_Com_Type_Name $_applicant.'電力申請会社名称' 

  # TODO: 派遣が含まれている場合、$_applicant.'管理会社'のみでよいかもしれない。要検証。
  if ($manegement_com.Contains("派遣")) { return $manegement_com }
  if ($manegement_com.Contains("ＴＲＳ―")) { return $manegement_com }
  if ($application_com -eq $employment_com) { return $application_com }
  return  ($application_com + " / " + $employment_com)
}

function private:fn_Extract_Reserved_Targets {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject[]][ref]$_reserved_persons
  )
  [PSCustomObject[]]$list = foreach ($_person in $_reserved_persons) {  
    [PSCustomObject]@{
      "漢字氏名"   = $_person."漢字氏名" 
      "所属企業名"  = $_person."登録_申請会社"
      "カタカナ氏名" = $_person. "カタカナ氏名"
      "中央登録番号" = $_person."中央登録番号"
    }
  }
  return $list
}

function private:fn_Extract_Registered_Targets {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject[]][ref]$_registered_persons
  )
  [PSCustomObject[]]$list = foreach ($_person in $_registered_persons) {  
    [PSCustomObject]@{
      "漢字氏名"   = fn_Combined_Name $_person."氏名（姓）" $_person."氏名（名）" 
      "所属企業名"  = fn_Application_Company_Names $_person
      "カタカナ氏名" = $_person."氏名（カナ）"
      "中央登録番号" = $_person."中登番号"
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

$reserved_info_Path = (${HOME} + $config.reserved_info.path)
[PSCustomObject[]]$reserved_info_list = fn_Read $reserved_info_Path

# 申請用オブジェクト生成 --------------------------------------
$search_key = $config.reserved_info.search_field
[PSCustomObject[]]$applicants = fn_SearchTargets ([ref]$reserved_info_list) $_regist_nums $search_key
#$applicants | Format-Table

if ($null -ne $applicants) {
  $script:formated_info_list = fn_Extract_Reserved_Targets ([ref]$applicants)
}
if ($null -eq $applicants) { 
  $private:_serch_key = $config.registerers_info.search_field
  $registerers_path = (${home} + $config.registerers_info.path)
  [PSCustomObject[]]$source_list = fn_Read $registerers_path
  [PSCustomObject[]]$script:registered_persons = fn_SearchTargets ([ref]$source_list) $_regist_nums $_serch_key
  $script:formated_info_list = fn_Extract_Registered_Targets ([ref]$registered_persons)
}


$kinds = $_Kinds.ToCharArray()

# $kinds[0] は ()で囲まなければアクセスできない。
$header = $config.($kinds[0]).address_table.psobject.properties.Name

[PSCustomObject[]]$applications = foreach ($_ in $formated_info_list) { $_ | Select-Object -Property $header } 
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

