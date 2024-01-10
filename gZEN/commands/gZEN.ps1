Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"



function script:fn_Init {
  Param(
    [Parameter(Mandatory = $True)]
    [ValidatePattern('\.json$')]
    [String]$_path
  )
  $json = Get-Content -Path $_path | ConvertFrom-Json
  return $json
}

function script:fn_Read_To_Array {
  Param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('\.txt$')]$txt_path,
    [String]$encode = "Default"
  )
  return Get-Content -Encoding $encode $txt_path
}

function script:fn_Find_Files {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$Folder,
    [Parameter(Mandatory = $True, Position = 1)]
    [String]$TargetName
  )
  $file_path_list = (Get-childItem -Path $Folder -File -Include $TargetName).fullname
  return $file_path_list
}

function script:fn_Bind_As_CSV {
  Param(
    [Object[]]$header,
    [Object[]]$values,
    [String]$delimiter = ','
  )

  [PSCustomObject[]]$csv_object = $values | ConvertFrom-Csv -Header $header.Split($delimiter)
  return $csv_object
}

function script:fn_Export_CSV {
  Param(
    [Parameter(Mandatory = $true, Position = 0)][Object[]]$_csv_obj,
    [Parameter(Mandatory = $true, Position = 1)][String]$_path,
    [String]$_delimiter = ',',
    [String]$_encode = "UTF8"# Default ではブラウザで参照すると文字化けする。
  )
  $_csv_obj | Export-Csv -NotypeInformation -Path $_path -Delimiter $_delimiter -Encoding $_encode -Force
}
function script:notifycation {
  Param(
    [String]$title,
    [String]$message
  )
  Add-Type -AssemblyName System.Windows.Forms
  $MUTEX_NAME = "Global\mutex" #多重起動チェック用

  try {
    $mutex = New-Object System.Threading.Mutex($False, $MUTEX_NAME)
    #多重起動チェック
    if ($mutex.WaitOne(0, $False)) {
      $notify_icon = New-Object Windows.Forms.NotifyIcon
      #$ApplicationContext = New-Object System.Windows.Forms.ApplicationContext

      $notify_icon.Icon = [Drawing.SystemIcons]::Application
      $notify_icon.Visible = $True
      # 通知用 のアイコン情報
      #$notify_icon.BalloonTipIcon = [Windows.Forms.ToolTipIcon]::Info

      $notify_icon.BalloonTipText = "$title :  $message"
      $notify_icon.ShowBalloonTip(1)

      # $_second 秒待機して通知を非表示にする。
      $notify_icon.Visible = $False
    }
  }
  finally {
    $notify_icon.Dispose()
    $mutex.ReleaseMutex()
    $mutex.Close()
    $mutex.Dispose()
    exit
  }
}

function script:fn_Combined_Name {
  Param(
    [Parameter(Mandatory = $true, Position = 0)][String]$first_name,
    [Parameter(Mandatory = $true, Position = 1)][String]$last_name,
    [String]$delimiter = '　'#デフォルト引数 呼び出し側で -delimiter を指定すること
  )

  $sb = New-Object System.Text.StringBuilder
  #副作用処理  StringBuilderならちょっと速いらしい。要素数が少ないから意味ないかも。
  @($first_name, $delimiter , $last_name) | ForEach-Object { [void] $sb.Append($_) }
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

function script:fn_Post_Code {
  Param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern("^\d{7}")][String]$arg
  )
  return Write-Output("{0:000-0000}" -f [Int]$arg)
}


function script:fn_To_Wide {
  Param(
    [Parameter(Mandatory = $True)][String]$half_string
  )
  Add-Type -AssemblyName "Microsoft.VisualBasic"
  [Microsoft.VisualBasic.Strings]::StrConv($half_string, [Microsoft.VisualBasic.VbStrConv]::Wide)
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

function script:fn_shape_values {
  Param(
    [PSObject[]]$_arg,
    [String[]]$_list
  )


  
  #プロパティの追加
  $new_obj = $_arg | Select-Object *, @{
    Name       = 'カタカナ氏名';
    Expression = { fn_Combined_Name $_.'カナ氏名（姓）' $_.'カナ氏名（名）' }
  }
  $new_obj = $new_obj | Select-Object *, @{
    Name       = '漢字氏名';
    Expression = { fn_Combined_Name $_.'漢字氏名（姓）' $_.'漢字氏名（名）' }
  }

  foreach ($_ in $new_obj) {
    # 郵便番号変換処理は不要かも
    $_.'現住所（住民票）郵便番号' = fn_Post_Code $_.'現住所（住民票）郵便番号'
    $_.'現住所（現在住んでいる）郵便番号' = fn_Post_Code $_.'現住所（現在住んでいる）郵便番号'
    $_.'現住所（住民票）住所' = fn_To_Wide $_.'現住所（住民票）住所'
    $_.'所属企業番号' = $_.'所属企業番号' -replace "^0", "T"
  }

  $coms_path = "${HOME}\Downloads\TEMP\Applicate_Coms.json"
  $coms = Get-Content -Path $coms_path | ConvertFrom-Json
  write-host $coms
  
  $new_obj = $new_obj | Select-Object *, @{
    Name       = '所属企業名';
    Expression = { $coms.($_.'所属企業番号') }
  }
  $new_obj = $new_obj | Select-Object *, @{
    Name       = '登録時申請会社';
    Expression = {
      $private:shorten_1 = fn_Shorten_Com_Type_Name $_.'所属企業名'
      $private:shorten_2 = fn_Shorten_Com_Type_Name $_.'雇用企業名称（漢字）'
      fn_Application_Company_Names $shorten_1 $shorten_2
    }
  }

  # 出力プロパティを任意の文字列[]で取得する。
  $selected_obj = $new_obj | Select-Object -Property $_list
  return $selected_obj
}


<#
try {
#>
# 設定読み込み

[PSCustomObject]$config = fn_Init ".\config\gzen.json"
#$config | Format-List

$header = fn_Read_To_Array $config.orign_header

# *事前申請*.txt はSHIFT-JISでエンコードされている。
$temps = $config.temp_folder
#$config.gZEN_targets
$values_filenames = fn_Find_Files "${HOME}$temps\*" $config.gZEN_targets
#$values_filenames

$values = foreach ($_ in $values_filenames) {
  fn_Read_To_Array $_
}
#$header.length
#$values.length

#テキストファイルを読み込み、ヘッダーをつけてCSVファイルを生成。
[PSCustomObject[]]$csv_obj = fn_Bind_As_CSV $header $values

# 必要な項目の情報を抽出する
$sorted_list = fn_Read_To_Array $config.sorted_header
$new_csv_obj = fn_shape_values $csv_obj $sorted_list
#$new_csv_obj| Format-List

$destination = (${HOME} + $config.export_path)
#$destination  
if (!(Test-Path $destination)) {
  # 空ファイルを作る
  New-Item -Path $destination -ItemType File -Force
}

#CSVファイルを出力
fn_Export_CSV $new_csv_obj $destination
  
#clipboardに出力する
$private:comma = ','
$private:tab = "`t"
$plain_text = Get-Content -Path $destination -Encoding UTF8
$formatted_text = $plain_text.Replace('"', '')
$formatted_text.Replace($comma, $tab) | Set-Clipboard
  

#ftの親パスを取得
$private:app_path = Convert-Path .
$private:cpy_path = ($app_path + $config.cpy_path_monolith)
#Test-Path($cpy_path)
$private:command_name = $config.command_name

# EDGEブラウザ起動し、cpyを起動する。
Start-Process msedge $cpy_path

# 通知を表示
notifycation $command_name ("🐈.,💩💩,,.  💩,  🌲🏡  Ver " + $config.version)
  
