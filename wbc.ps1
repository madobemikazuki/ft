Param(
  [Parameter(Mandatory = $True, Position = 0) ]
  [ValidatePattern("r|c")]$_task,

  [Parameter(Mandatory = $True, Position = 1)]
  [ValidatePattern("^\d{8}")][String]$_date,

  [Parameter(Mandatory = $True, Position = 2)]
  [ValidateCount(1, 4)]
  [String[]]$_regists
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
  $output_file_path = $io_object.output + $io_object.headname + $io_object.class + "_" + $_full_names + $io_object.extension
  # 中身の気になるあなたに
  #Write-Output $values
  . .\excel\read.ps1
  try {
    
    $book = read_xlsx_nonVisible $io_object.template
    <# Worksheets.Item(シート名) で指定したシートを開きます。
      注意点として、ExcelはSJISなので、シート名が日本語のときは、
      PowerShellのファイルはSJISにして実行する必要があります。
      PowerShellのファイルを UTF-8 で保存すると、日本語のシート名が検索できないので、
      代わりに .Worksheets.Item(シート番号) とする方法もあります。
    #>
    $sheet = $book.Worksheets.Item(1)

    #mappingする
    foreach ($_ in $values) {
      $sheet.Cells.Item($_.point_x, $_.point_y) = $_.value
    }
    
    # 関数へ切り出す
    # Printer名をどうにかしないと
    #$printer = Get-WmiObject Win32_Printer | Where-Object Name -eq "Wi-Fi Direct DCP-J526N"
    #$printer.SetDefaultPrinter()
    #Set-PrintConfiguration $printer.name -Color $False

    # excel ファイル自体に両面印刷設定されていればOK
    #Set-PrintConfiguration -PrinterName "B_J526N_USB" -DuplexingMode TwoSidedShortEdge

    # 関数へ切り出す
    $start_page = 1
    $end_page = 2
    $number_of_copies = 1


    # プリントアウトする
    $book.PrintOut.Invoke(@($start_page, $end_page, $number_of_copies))

    #別ファイルに書き
    . .\excel\write.ps1
    write_xslx $book $output_file_path
    
  }
  catch [exception] {
    Write-Output "😢😢😢エラーをよく読んでね。"
    $error[0].ToString()
  }
  finally {
    @($sheet, $book) | ForEach-Object {
      if ($_ -ne $null) {
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($_)
      }
    }
    quit_excel
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


function script:fn_create_io_path_object {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$_template_file_path,
    [Parameter(Mandatory = $True, Position = 1)]
    [String]$_output_folder_path,
    [Parameter(Mandatory = $True, Position = 2)]
    [ValidatePattern("登録|解除")]
    [string]$_class
  )

  $io_object = [PSCustomObject]@{
    template  = $_template_file_path
    output    = $_output_folder_path
    class     = $_class
    headname  = "WBC_"
    extension = ".xlsx"
  }
  return $io_object
}


#申請日をフォーマット定義
$application_date = fn_apply_date $_date


# 登録者を読み込み
$applicant_list = . .\ft_core\io\csv\read_applicants_fromT.ps1


#今回の申請者を抽出する
#ここで例外をキャッチする もっとましな書き方が望ましい。
[PSCustomObject[]]$applicants = . .\ft_core\search_applicants.ps1 $applicant_list $_regists
if (!($applicants.length -eq $_regists.length)) {
  Write-Host '貴様の入力した中登番号は登録状況リストには存在しない。'
  # 存在しない中央登録番号を入力するとapplicants.lengthが小さくなる。
  throw
}


#転記先のアドレスを定義
$address_table = Import-Csv ${HOME}\Downloads\config\wbc_address_table.csv -Encoding utf8
$HEADER = $address_table[0].psobject.Properties.Name

$io_object = $Null

# 登録モード
if ($_task -eq "r") {
  Write-Host "${_task} : 登録だね"
  $private:wbc_r_xlsx_file = "${HOME}\Downloads\from_T\登録\WBC受検用紙_登録_原紙.xlsx"
  $private:entered_wbc_r_folder = "${HOME}\Downloads\output\登録\WBC受検用紙\"
  $io_object = fn_create_io_path_object $wbc_r_xlsx_file $entered_wbc_r_folder "登録"
}

# 解除モード
if ($_task -eq "c") {
  $private:wbc_c_xlsx_file = "${HOME}\Downloads\from_T\解除\WBC受検用紙_解除_原紙.xlsx"
  $private:entered_wbc_c_folder = "${HOME}\Downloads\output\解除\WBC受検用紙\"
  Write-Host "${_task} : 解除だね"
  $io_object = fn_create_io_path_object $wbc_c_xlsx_file $entered_wbc_c_folder "解除"
}

[PSCustomObject[]]$applicants_info = fn_extract $application_date $applicants $HEADER
#$applicants_info
#Write-Host '245'
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
$for_posting | format-table

fn_transcription $for_posting $io_object $applicant_names



exit 0