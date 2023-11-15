Set-StrictMode -Version 3.0

# 必要な情報は設定ファイルとして読み込む init()が必要

$gzen_header = ".\data\header\gZEN_header_ANSI.txt"
$selected_list = ".\data\header\gZen_select_items_ANSI.txt"

$export_folder = "${HOME}\Downloads"
$target_name = "*事前申*.txt"
$export_file_path = "$export_folder\export.csv" 

$comma = ','
$tab = "`t"

#ftの親パスを取得
$app_path = Get-Location | Split-Path
$cpy_path = "$app_path\ft\web_apps\cpy_ft\cpy_ft.html"
$command_name = "gZEN"

function notifycation {
  Param(
    [String]$title,
    [String]$message
  )
  . .\commands\utils\notify.ps1 
  notify_balloon $title $message
}

function shape_values {
  Param(
    [PSObject[]]$arg,
    [String[]]$list
  )
  . .\commands\ft_core\combined_name.ps1
  . .\commands\utils\util_format.ps1
  #プロパティの追加
  $new_obj = $arg | Select-Object *, @{
    Name       = 'カタカナ氏名';
    Expression = { combined_name $_.'カナ氏名（姓）' $_.'カナ氏名（名）' }
  }
  $new_obj = $new_obj | Select-Object *, @{
    Name       = '漢字氏名';
    Expression = { combined_name $_.'漢字氏名（姓）' $_.'漢字氏名（名）' }
  }

  . .\commands\ft_core\hanzen.ps1

  foreach ($_ in $new_obj) {
    $_.'現住所（住民票）郵便番号' = post_code $_.'現住所（住民票）郵便番号'
    $_.'現住所（現在住んでいる）郵便番号' = post_code $_.'現住所（現在住んでいる）郵便番号'
    $_.'現住所（住民票）住所' = to_wide $_.'現住所（住民票）住所'
    $_.psobject.properties.remove('カナ氏名（姓）')
    $_.psobject.properties.remove('カナ氏名（名）')
  }
  # 出力プロパティを任意の文字列[]で取得する。
  $selected_obj = $new_obj | Select-Object -Property $list
  return $selected_obj
}



try {
  . .\commands\utils\util_txt.ps1
  $header = read_to_array $gzen_header
  # *事前申請*.txt はSHIFT-JISでエンコードされている。
  $values_filename = find_file_name $export_folder $target_name
  $values = read_to_array "$export_folder\$values_filename"

  #テキストファイルを読み込み、ヘッダーをつけてCSVファイルを生成。
  . .\commands\utils\util_csv.ps1
  $csv_obj = bind_as_csv $header $values

  # 必要な項目の情報を抽出する
  $list = read_to_array $selected_list
  $selected_csv_obj = $csv_obj | Select-Object -Property $list

  $output_list = read_to_array ".\data\header\gZen_output_items_ANSI.txt"
  $new_csv_obj = shape_values $selected_csv_obj $output_list
  #CSVファイルを出力
  export_csv $new_csv_obj $export_file_path $comma

  #clipboardに出力する
  $plain_text = Get-Content -Path $export_file_path -Encoding UTF8
  $formatted_text = $plain_text.Replace('"', '')
  $formatted_text.Replace($comma, $tab) | Set-Clipboard

  # $values＿file_name の削除
  #Remove-Item -Path $values_filename

  # EDGEブラウザ起動し、cpyを起動する。
  Start-Process msedge $cpy_path


  # 通知を表示
  notifycation $command_name "🐈.,💩💩,,.  💩,  🌲🏡"
}
catch {
  Write-Host "エラー発生 :: $($_.Exception.Message)"
  notifycation "エラー発生 ::" "$($_.Exception.Message)"
}
