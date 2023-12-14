Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"



function fn_init {
  . .\commands\ft_core\io\read_json.ps1 .\commands\config\gZEN.json
}

function notifycation {
  Param(
    [String]$title,
    [String]$message
  )
  . .\commands\utils\notify_balloon.ps1 $title $message
}

function fn_shape_values {
  Param(
    [PSObject[]]$arg,
    [String[]]$list
  )

  #プロパティの追加
  $new_obj = $arg | Select-Object *, @{
    Name       = 'カタカナ氏名';
    Expression = { . .\commands\ft_core\combined_name.ps1 $_.'カナ氏名（姓）' $_.'カナ氏名（名）' }
  }
  $new_obj = $new_obj | Select-Object *, @{
    Name       = '漢字氏名';
    Expression = { . .\commands\ft_core\combined_name.ps1 $_.'漢字氏名（姓）' $_.'漢字氏名（名）' }
  }

  foreach ($_ in $new_obj) {
    # 郵便番号変換処理は不要かも
    $_.'現住所（住民票）郵便番号' = . .\commands\utils\post_code.ps1 $_.'現住所（住民票）郵便番号'
    $_.'現住所（現在住んでいる）郵便番号' = . .\commands\utils\post_code.ps1 $_.'現住所（現在住んでいる）郵便番号'
    $_.'現住所（住民票）住所' = . .\commands\utils\hanzen.ps1 wide $_.'現住所（住民票）住所'
    #$_.psobject.properties.remove('カナ氏名（姓）')
    #$_.psobject.properties.remove('カナ氏名（名）')
  }
  # 出力プロパティを任意の文字列[]で取得する。
  $selected_obj = $new_obj | Select-Object -Property $list
  return $selected_obj
}


<#
try {
#>
  # 設定読み込み
  [PSCustomObject]$config = fn_init
  #$config | Format-List

  $header = . .\commands\ft_core\io\read_to_array.ps1 $config.orign_header
 
  # *事前申請*.txt はSHIFT-JISでエンコードされている。
  $temps = $config.temp_folder
  #$config.gZEN_targets
  $values_filenames = . .\commands\ft_core\io\find_files.ps1 "${HOME}$temps\*" $config.gZEN_targets
  #$values_filenames

  $values = foreach ($_ in $values_filenames) {
    . .\commands\ft_core\io\read_to_array.ps1 $_
  }
  #$header.length
  #$values.length

  #テキストファイルを読み込み、ヘッダーをつけてCSVファイルを生成。
  [PSCustomObject[]]$csv_obj = . .\commands\ft_core\io\bind_as_csv.ps1 $header $values

  # 必要な項目の情報を抽出する
  $sorted_list = . .\commands\ft_core\io\read_to_array.ps1 $config.sorted_header
  $new_csv_obj = fn_shape_values $csv_obj $sorted_list
  #$new_csv_obj| Format-List

  $destination = (${HOME} + $config.export_path)
  #$destination  
  if (!(Test-Path $destination)) {
    # 空ファイルを作る
    New-Item -Path $destination -ItemType File -Force
  }

  #CSVファイルを出力
  . .\commands\ft_core\io\export_csv.ps1 $new_csv_obj $destination
  
  #clipboardに出力する
  $private:comma = ','
  $private:tab = "`t"
  $plain_text = Get-Content -Path $destination -Encoding UTF8
  $formatted_text = $plain_text.Replace('"', '')
  $formatted_text.Replace($comma, $tab) | Set-Clipboard
  

  #ftの親パスを取得
  $private:app_path = Convert-Path .
  $private:cpy_path = ($app_path + $config.cpy_path)
  #Test-Path($cpy_path)
  $private:command_name = $config.command_name

  # EDGEブラウザ起動し、cpyを起動する。
  Start-Process msedge $cpy_path

  # 通知を表示
  notifycation $command_name "🐈.,💩💩,,.  💩,  🌲🏡"
  

<#}
catch {
  Write-Host "エラー発生 :: $($_.Exception.Message)"
  #notifycation "エラー発生 ::" "$($_.Exception.Message)"
}

#>
