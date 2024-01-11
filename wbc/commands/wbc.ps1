Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [ValidatePattern("r|c")]$Task,
  [Parameter(Mandatory = $True, Position = 1)]
  [ValidatePattern("^\d{8}")][String]$_date,
  [Parameter(Mandatory = $True, Position = 2)]
  [ValidatePattern("^\d{2}\b\-\b\d{6}$")]
  [ValidateCount(1, 4)][String[]]$_central_nums
)

<#
  動作チェック

  .\commands\wbc.ps1 r のときは
  \Downloads\TEMP\gZEN_exported.csv から "中央登録番号" を選択すること

  .\commands\wbc.ps1 c のときは
  \Downloads\TEMP\登録者管理リスト.csv から "中登番号"を選択すること

#>

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"




function script:fn_Apply_Date {
  Param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern("^\d{8}")][String]$_arg
  )
  try {
    $date = [DateTime]::ParseExact($_arg, 'yyyyMMdd', $null)
    if (fn_Is_Future $date ) {
      return fn_Excel_Hell_Format $date
    }
    else {
      Throw
    }
    <#
  .SYNOPSIS
  貴様が明日以降の日付を入力したのならWBC受検用の日付フォーマット文字列を返してやる。
  #>
  }
  catch {
    Throw "今日より明日だ。覚えておけ。"
  }
  
}

function fn_Is_Future {
  Param
  (
    [Parameter(Mandatory)]
    [DateTime]$_future
  )
  $today = Get-Date
  $_future -gt $today
}

function fn_Excel_Hell_Format {
  Param(
    [Parameter(Mandatory)]
    [DateTime]$_future
  )
  return $_future.ToString("yyyy年　　MM月　　dd日");
}


function script:fn_Combined_Name {
  Param(
    [Parameter(Mandatory = $true, Position=0)][String]$_first_name,
    [Parameter(Mandatory = $true, Position=1)][String]$_last_name,
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
    return fn_Combined_Name $_managemanet_com_name $_employer_name  -delimiter " / "
  }
}

function script:fn_Shorten_Com_Type_Name {
  Param(
    [Parameter(Mandatory = $True)]
    [String]$_corporate_name
  )
  switch ($_corporate_name) {
    { $_.Contains('株式会社') } { return $_.Replace('株式会社', '（株）') }
    { $_.Contains('有限会社') } { return $_.Replace('有限会社', '（有）') }
  }
}

function script:fn_Search_Target {
  Param(
    [Parameter(Mandatory = $True, Position = 0) ]
    [PSCustomObject[]][ref]$_psobject_array,
    [Parameter(Mandatory = $True, Position = 1) ]
    [String]$_target,
    [Parameter(Mandatory = $True, Position = 2)]
    [String]$_flag
  )
  # $_flag には'中央登録番号' もしくは '中登番号' が入る
  [PSCustomObject[]]$result = $_psobject_array | Where-Object { $_.$_flag -eq $target } 
  return $result
}

function fn_PSObjList_Filter {
  Param(
    [Parameter(Mandatory = $True, Position = 0) ]
    [PSCustomObject[]][ref]$_object_array,
    [Parameter(Mandatory = $True, Position = 1) ]
    [String[]]$_targets
  )
  $collection = foreach ($object in ($_object_array)) {
    $object | Select-Object -Property $_targets
  }
  return $collection
}

function fn_Generate_WBC_Config_Object {
  Param(
    [Parameter(Mandatory = $True)]
    [ValidatePattern('r|c')]$_Task,
    [Parameter(Mandatory = $True)]
    [String]$_config_path
  )
  $private:config = Get-Content -Path $_config_path | ConvertFrom-Json

  $private:obj = [PSCustomObject]@{
    task         = $config.$_Task.task
    source_csv   = $config.$_Task.source_csv
    command_name = $config.command_name
    extension    = $config.extension
    template     = (${HOME} + $config.$_Task.tamplate_file)
    output       = (${HOME} + $config.$_Task.output_folder)
    printing     = $config.printing
  }
  return $obj
}


function fn_Generate_WBC_Output_Path {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject][ref]$_wbc_config,
    [Parameter(Mandatory = $True, Position = 1)]
    [String]$_names
  )

  $output_file_path = @(
    $_wbc_config.output,
    $_wbc_config.task,
    $_wbc_config.command_name,
    '_',
    $_names,
    $_wbc_config.extension
  ) -Join ''
  return $output_file_path
}

function fn_Posting_Format {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject[]][ref]$_applicants,
    [Parameter(Mandatory = $True, Position = 1)]
    [String[]][ref]$_header,
    [Parameter(Mandatory = $True, position = 2)]
    [PSCustomObject[]][ref]$io_object
  )
  
  $formated_obj = foreach ($applicant in $_applicants) {
    $index = $_applicants.indexOf($applicant)
    if ($applicant.psobject.Properties.value.count -eq $_header.length) {
      $position = $io_object.printing.address_table[$index]
      fn_Map_Address $_header $position $applicant 
    }
  }
  return $formated_obj
}

function fn_Map_Address {
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

function private:fn_Transcription {
  Param(
    [Parameter(Mandatory = $true, Position = 0)]
    [PSCustomObject[]]$_posting_object,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject]$_config,
    [Parameter(Mandatory = $True, position = 2)]
    [String]$_export_path
  )
  try {
    # Measure-Command でブロック内の実行完了時間を測定できる。
    $time = Measure-Command {
      $excel = New-Object -ComObject Excel.Application
      #.Visible = $false でExcelを表示しないで処理を実行できる。
      $excel.Visible = $False
      # 上書き保存時に表示されるアラートなどを非表示にする
      $excel.DisplayAlerts = $False
      # リンクの更新方法が 0 の場合は何もしない。
      #.Workbooks.Open(ファイル名, リンクの更新方法, 読み取り専用) でExcelを開きます。
      $script:book = $excel.Workbooks.Open($_config.template, 0, $true)
    }
    Write-host $time.TotalSeconds.ToString("F2")"秒 : Excelの起動が完了するまでの経過時間"
    <# Worksheets.Item(シート名) で指定したシートを開きます。
      注意点として、ExcelはSJISなので、シート名が日本語のときは、
      PowerShellのファイルはSJISにして実行する必要があります。
      PowerShellのファイルを UTF-8 で保存すると、日本語のシート名が検索できないので、
      代わりに .Worksheets.Item(シート番号) とする方法もあります。
    #>
    $sheet = $book.Worksheets.Item($_config.printing.sheet_page)

    foreach ($_ in $_posting_object) {
      $sheet.Cells.Item($_.point_x, $_.point_y) = $_.value
    }


    <#
    # プリントアウトする
    $default = Get-WmiObject Win32_Printer | Where-Object default
    $print_config = $_config.printing

    #今から使うプリンタを設定  プリンタ名が指定されないと例外が発生しスクリプトは止まる。
    $printer = Get-WmiObject Win32_Printer | Where-Object name -eq $print_config.printer_name
    $printer.SetDefaultPrinter()
    #Set-PrintConfiguration -PrinterName $printer.name -Color $print_config.color
    
    $start = [int16]$print_config.start_page
    $end = [int16]$print_config.end_page
    $copies = [int16]$print_config.number_of_copies
    
    # プリントアウトする
    $book.PrintOut.Invoke(@($start, $end, $copies))
    #プリンタ設定をプリントアウト前の設定に戻す
    $default.SetDefaultPrinter()
#>


    # 空ファイルを作成し、そこに出力
    New-Item $_export_path -type file -Force
    $book.SaveAs($_export_path)
    #$values | Format-Table   
    Write-Output "👍👍👍  出力先 : $_export_path"    
    $book.Close()
  }
  catch [exception] {
    Write-Output "😢😢😢エラーをよく読んでね。"
    $error[0].ToString()
    Write-Output $_
  }
  finally {
    @($sheet, $book) | ForEach-Object {
      if ($_ -ne $null) {
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($_)
      }
    }
    $excel.Quit()
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
  }
}

# TEPCOに提出した登録の事前申請書 から情報を取得し、整形して返す
function fn_Registration_Format {
  Param(
    [Parameter(Mandatory = $True, Position = 0) ]
    [String]$_application_date,
    [Parameter(Mandatory = $True, Position = 1)]
    [ValidateCount(1, 4)][String[]]$_applicants,
    [Parameter(Mandatory = $True, Position = 2)]
    [String]$_source
  )
  Write-Host '登録モードだよ'

  #事前申請書のCSVから情報を取得
  $csv_obj = Import-Csv -Path "${HOME}$_source" -Encoding UTF8
  #$csv_obj | Select-Object {$_.'中央登録番号'} |Format-Table

  #事前申請書のヘッダーに基づく抽出対象
  $private:extract_list = @(
    '中央登録番号',
    '個人番号'
    '漢字氏名',
    '登録時申請会社'
  )

  $targets = foreach ($target in $_applicants) {
    fn_Search_Target ([ref]$csv_obj) $target ($extract_list[0])
  }
  
  [PSCustomObject[]]$extracted_targets_info = fn_PSObjList_Filter ([ref]$targets) $extract_list
  [PSCustomObject[]]$applicants = foreach ($_ in $extracted_targets_info) {
    
    [PSCustomObject] @{
      $script:wbc_application_field[0] = $_application_date
      $script:wbc_application_field[1] = $_.'個人番号'
      $script:wbc_application_field[2] = $_.'登録時申請会社'
      $script:wbc_application_field[3] = $_.'漢字氏名'
    }
  }

  return $applicants
}


# 登録状況リスト と wid_gr.jsonから情報を取得し、整形して返す
function fn_Cancellation_Format {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$_application_date,
    [Parameter(Mandatory = $True, Position = 1)]
    [ValidatePattern("^\d{2}\b\-\b\d{6}$")]
    [String[]]$_regist_nums,
    [Parameter(Mandatory = $True, Position = 2)]
    [String]$_source
  )

  Write-Host '解除モードだよ'
  #登録者管理リストのヘッダーに基づく抽出対象
  $extract_list = @(
    '中登番号',
    '作業者証番号',
    '氏名（姓）',
    '氏名（名）',
    '電力申請会社名称',
    '雇用名称'
  )

  # これは読み込めてなさそう？
  # config に含めること?
  $registed_list = Import-Csv -Path "${HOME}$_source" -Encoding UTF8

  $targets = foreach ($target in $_regist_nums) {
    fn_Search_Target ([ref]$registed_list) $target $extract_list[0]
  }

  [PSCustomObject[]]$extracted_targets_info = fn_PSObjList_Filter ([ref]$targets) $extract_list
  [PSCustomObject[]]$private:applicants = foreach ($_ in $extracted_targets_info) {

    $shorten_1 = $_.($extract_list[4])
    $shorten_2 = $_.($extract_list[5])
    #$shorten_1 = fn_Shorten_Com_Type_Name $_.($extract_list[4])
    #$shorten_2 = fn_Shorten_Com_Type_Name $_.($extract_list[5])


    [PSCustomObject] @{
      $script:wbc_application_field[0] = $_application_date
      $script:wbc_application_field[1] = $_.($extract_list[1])
      
      # ◆未実装  所属企業番号から '申請会社名' を取得すること 
      $script:wbc_application_field[2] = fn_Application_Company_Names $shorten_1 $shorten_2
      $script:wbc_application_field[3] = fn_Combined_Name $_.'氏名（姓）' $_.'氏名（名）'
    }
  }
  return $applicants
}

function private:fn_display {
  Param(
    [Parameter(Mandatory = $True)]
    [String]$_message
  )
  return  $_message + '楽しんでね！'
}


<#
  ここからコマンドの実行内容
#>

$script:wbc_application_field = @(
  "申請日",
  "東電作業者証番号",
  "会社名",
  "氏名"
)

$private:application_date = fn_Apply_Date $_date
[PSCustomObject[]]$private:applicants = @{}

$private:config_path = ".\config\wbc.json"
$private:config_object = fn_Generate_WBC_Config_Object $Task $config_path

#ここで必要な情報を集約整形して格納する。
switch ($Task) {
  r { $applicants = fn_Registration_Format $application_date $_central_nums $config_object.source_csv }
  c { $applicants = fn_Cancellation_Format $application_date $_central_nums $config_object.source_csv }
}


#$private:application_date
$applicants | Format-Table
#$private:config_object = fn_Generate_WBC_Config_Object $Task


$name_list = fn_PSObjList_Filter ([ref]$applicants) @($wbc_application_field[3])
$list = foreach ($_ in $name_list) { $_.($wbc_application_field[3]) }
$one_line_names = $list -join '_'
$export_path = fn_Generate_WBC_Output_Path ([ref]$private:config_object) $one_line_names


[PSCustomObject[]]$posting_object = fn_Posting_Format ([ref]$applicants) ([ref]$wbc_application_field) ([ref]$config_object)
#$posting_object | Format-Table

# transcription.ps1 にしたいなぁ。
fn_Transcription $posting_object $config_object $export_path

exit 0
