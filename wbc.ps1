Param(
  [Parameter(Mandatory = $True, Position = 0) ]
  [ValidatePattern("r|c")]$_class,

  [Parameter(Mandatory = $True, Position = 1)]
  [ValidatePattern("^\d{8}")][String]$_date,

  [Parameter(Mandatory = $True, Position = 2)]
  [ValidateCount(1, 4)]
  [String[]]$_applicants
)


Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"


function private:fn_transcription {
  Param(
    [Parameter(Mandatory = $true, Position = 0)]
    [PSOBject[]]$values,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject]$io_object,
    [Parameter(Mandatory = $True, position = 2)]
    [String]$_full_names
  )
  
  $output_file_path = $io_object.output + $io_object.task + $io_object.headname + "_" + $_full_names + $io_object.extension
  #Write-Host "出力先: $output_file_path"
  $target_sheet_page = 1

  try {
    # Measure-Command でブロック内の実行完了時間を測定できる。
    $time = Measure-Command {
      $excel = New-Object -ComObject Excel.Application
      #.Visible = $false でExcelを表示しないで処理を実行できる。
      $excel.Visible = $False
      # 上書き保存時に表示されるアラートなどを非表示にする
      $excel.DisplayAlerts = $False
      # リンクの更新方法が 0 の場合は何もしない。
      #.Workbooks.Open(ファイル名, リンクの更新方法, 読み取り専用) でExcelを開きます1。
      $script:book = $excel.Workbooks.Open($io_object.template, 0, $true)
    }
    Write-host $time.TotalSeconds.ToString("F2")"秒 : Excelの起動が完了するまでの経過時間"
    <# Worksheets.Item(シート名) で指定したシートを開きます。
      注意点として、ExcelはSJISなので、シート名が日本語のときは、
      PowerShellのファイルはSJISにして実行する必要があります。
      PowerShellのファイルを UTF-8 で保存すると、日本語のシート名が検索できないので、
      代わりに .Worksheets.Item(シート番号) とする方法もあります。
    #>
    $sheet = $book.Worksheets.Item($target_sheet_page)

    foreach ($_ in $values) {
      $sheet.Cells.Item($_.point_x, $_.point_y) = $_.value
    }


<#
    # プリントアウトする
    $default = Get-WmiObject Win32_Printer | Where-Object default
    $print_config = $io_object.printing

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


    # 新しいxlsx ファイルに出力
    $book.SaveAs($output_file_path)
    #$values | Format-Table   
    Write-Output "👍👍👍  出力先 : $output_file_path"    
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




function private:fn_obj_mapping {
  Param(
    [Parameter(Mandatory = $true)]
    [PSCustomObject]$applicant_info,
    [PSCustomObject]$position,
    [String[]]$header
  )
  . .\excel\map_address.ps1
  map_address $header $position $applicant_info
}


function private:fn_extract {
  Param(
    [Parameter(Mandatory = $true, Position = 0)]
    [String]$application_date,
    [Parameter(Mandatory = $true, Position = 1)]
    [ValidateCount(1, 4)][PSCustomObject[]]$target_applicants,
    [Parameter(Mandatory = $true, Position = 2)][String[]]$header
  )

  #wbc コマンド内で利用する変数
  $first_name = "氏名（姓）"
  $last_name = "氏名（名）"
  $company_name = "会社名"
  $employer_name = "雇用会社名"
  $tepco_on_site_license = "東電作業者証番号"

  #めっちゃ速くなった！
  #もっと速くすべきです。いちいちファイルアクセスが発生している。2023/10/22
  $applicant_info = foreach ($_ in $target_applicants) {
    [PSCustomObject] @{
      $header[0] = $application_date
      $header[1] = $_.$tepco_on_site_license
      $header[2] = fn_regist_company_names $_.$company_name, $_.$employer_name
      $header[3] = fn_regist_name @($_.$first_name, $_.$last_name)
    }
  }
  return $applicant_info
}


function script:fn_regist_name {
  Param(
    [Parameter(Mandatory = $true)]
    [Array]$names
  )
  . .\ft_core\combined_name.ps1
  combined_name $names[0] $names[1]
}

function script:fn_regist_company_names {
  Param(
    [Parameter(Mandatory = $true)]
    [Array]$names
  )
  . .\ft_core\your_company_names.ps1
  your_company_names $names[0] $names[1]
}


function private:fn_apply_date {
  Param(
    [Parameter(Mandatory = $true)]
    [String]$date
  )
  . .\ft_core\future_date.ps1
  future_date $date | excel_hell_format
  <#
  .SYNOPSIS
  貴様が明日以降の日付を入力したのならWBC受検用の日付フォーマット文字列を返してやる。
  #>
}


function wbc_example {
  [cmdletbinding()]
  Param()
  <#
  .SYNOPSIS
  WBC受検用紙をプリンタから出力するよ。
  .DESCRIPTION
  貴様らの用意したクソスぺPCでちまちま Excel ファイルにコピペしてプリントするが面倒だから作ってやったんだ。貴様はこうべを深々と下げて、私の知性と寛容さに感謝するのがよかろう。
  .EXAMPLE
  wbc r 20530203 00-123456,11-123456,22-123456,33-123456
  登録者は -r を指定すること。
  .EXAMPLE
  wbc c 20530203 00-123456,11-123456,22-123456,33-123456
  解除者は -c を指定すること。
  .OUTPUTS
  1: /downloads/output/登録/  もしくは  /downloads/output/削除
  .OUTPUTS
  2: 神の思し召しがあれば最寄りのプリンタから紙で出力されるよ。
  #>
  Write-Host ""
  Write-Host ""
  Write-Host ""
  Write-Host "🌞 コマンド入力例"
  Write-Host "💻> . .\wbc.ps1 r 20331010 96-498360,24-882388,63-035569,54-994059"
  Write-Host ""
  Write-Host ""
  Write-Host "出力先 : 📁 /downloads/output/登録/WBC/  もしくは  /downloads/output/解除/WBC/"
  Write-Host "出力先 : 📄 神の思し召しがあれば最寄りのプリンタから出力されるよ。"
  Write-Host ""
  Write-Host ""
  Write-Host ""
}


#申請日をフォーマット定義
$application_date = fn_apply_date $_date

# 登録者を読み込み
$private:registed_list = . .\ft_core\io\read_registed_people_fromT.ps1


#今回の申請者を抽出する
#ここで例外をキャッチする もっとましな書き方が望ましい。
[PSCustomObject[]]$applicants = . .\ft_core\search_applicants.ps1 $registed_list $_applicants
if (!($applicants.length -eq $_applicants.length)) {
  Write-Host '貴様の入力した中登番号は登録状況リストには存在しない。'
  # 存在しない中央登録番号を入力するとapplicants.lengthが小さくなる。
  throw
}


# 設定読み込み
$private:config = . .\ft_core\io\read_json.ps1 ".\config\wbc.json"
$private:io_object = [PSCustomObject]@{
  task      = $config.$_class.task
  headname  = $config.command_name
  extension = $config.extension
  template  = (${HOME} + $config.$_class.tamplate_file)
  output    = (${HOME} + $config.$_class.output_folder)
  printing  = $config.printing
}


#転記先のアドレスを定義
$address_table = Import-Csv -Path ($config.address_table_file) -Encoding utf8
$HEADER = $address_table[0].psobject.Properties.Name


[PSCustomObject[]]$applicants_info = fn_extract $application_date $applicants $HEADER
$private:full_name_list = foreach ($_ in $applicants_info) { $_."氏名" }
. .\ft_core\combined_name.ps1
$applicant_names = one_liner $full_name_list


[PSCustomObject[]]$for_posting = foreach ($applicant in $applicants_info) {
  $index = $applicants_info.indexOf($applicant)
  if ($applicant.psobject.Properties.value.count -eq $HEADER.length) {
    $position = $address_table[$index]
    fn_obj_mapping $applicant $position $HEADER
  }
}

#結果を出力しなくてもよい
#$for_posting | format-table

fn_transcription $for_posting $io_object $applicant_names

exit 0