Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"
<#
  \Downloads\PAN配下の *事前申請*.txt を読み込んで
  \config\csv_header\gZEN_header_ANSI.txtとともに
  申請会社情報 : Downloads\TEMP\Applicate_Coms.json
  を中央登録番号に基づきPSCustomObject[]に格納し、
  \config\csv_header\gZEN_sorted_ANSI.txt に指定したフィールドの順番で
  JSON形式とCSV方式に出力する。
#>



function script:fn_Read_JSON {
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
  $files = (Get-ChildItem -Path $Folder -File -Filter $TargetName).fullname
  return $files
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
    [String]$_encode = "utf8"# Default ではブラウザで参照すると文字化けする。
  )
  $_csv_obj | Export-Csv -NotypeInformation -Path $_path -Delimiter $_delimiter -Encoding $_encode -Force
}

function script:fn_Export_JSON {
  Param(
    [Parameter(Mandatory = $true, Position = 0)][Object[]]$_obj,
    [Parameter(Mandatory = $true, Position = 1)][String]$_path
  )
  $utf8_with_BOM = New-Object System.Text.UTF8Encoding $True
  [System.IO.File]::WriteAllLines($_path, (ConvertTo-Json $_obj), $utf8_with_BOM)
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

function script:fn_Sort_by_Array {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject[]]$_source_list,
    [Parameter(Mandatory = $True, Position = 1)]
    [String[]]$_sorted_field_list
  )
  # $_source_list より短い長さの $_sorted_fiel でもok
  $sorted_list = $_source_list | Select-Object -Property $_sorted_field_list
  return $sorted_list
}

function script:fn_Array_Filter {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject[]]$_source_list,
    [Parameter(Mandatory = $True, Position = 1)]
    [String]$_prop,
    [Parameter(Mandatory = $True, Position = 2)]
    [String]$_value
  )
  $new_list = $_source_list | Where-Object { $_.$_prop -eq $_value }
  return $new_list
}

function script:fn_Transform {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject[]]$_obj_list
  )
  $new_list = foreach ($_ in $_obj_list) {  
    $_ | Add-Member -NotePropertyMembers @{
      '所属企業番号' = $_.'所属企業番号' -replace "^0", "T"
      '現住所（住民票）住所' = fn_To_Wide $_.'現住所（住民票）住所'
      '漢字氏名' = fn_Combined_Name $_.'漢字氏名（姓）' $_.'漢字氏名（名）'
      'カタカナ氏名' = fn_Combined_Name $_.'カナ氏名（姓）' $_.'カナ氏名（名）'
      '現住所（住民票）郵便番号' = fn_Post_Code $_.'現住所（住民票）郵便番号'
      '現住所（現在住んでいる）郵便番号' = fn_Post_Code $_.'現住所（現在住んでいる）郵便番号'
    } -Force
  }
  return $new_list
}

function script:fn_Shape_Values {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSObject[]]$_arg,
    [Parameter(Mandatory = $True, Position = 1)]
    [String[]]$_list,
    [Parameter(Mandatory = $True, Position = 2)]
    [String]$_app_coms_path
  )

  #$new_obj = fn_Transform $_arg

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

  $coms_path = (${HOME} + $_app_coms_path)
  [PSCustomObject]$private:coms = Get-Content -Path $coms_path | ConvertFrom-Json
  [PSCustomObject]$app_coms = $coms.app_coms
  #Write-Host $app_coms
  
  $new_obj = $new_obj | Select-Object *, @{
    Name       = '所属企業名';
    Expression = { $app_coms.($_.'所属企業番号') }
  }

  foreach ($_ in $new_obj) {
    if ($app_coms.psobject.properties.name -notcontains $_.'所属企業番号') {
      Start-Process -FilePath notepad.exe -ArgumentList $coms_path
      #先に通知を表示すると別プロセスを起動できなくなる。
      notifycation "gZEN" (($_.'所属企業番号') + " : 申請企業名を書き足してください。")
      exit 0
    }
  }

  # 出力プロパティを任意の順序で取得する。
  return fn_Sort_by_Array $new_obj $_list
}



function script:fn_Move_To_Waste {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String[]]$_file_list,
    [Parameter(Mandatory = $True, Position = 1)]
    [String]$_waste_folder
  )
  foreach ($_file in $_file_list) {
    Move-Item -Path $_file -Destination $_waste_folder
  }
}


<#
--------------------- ここから実行内容 -----------------------
#>

[PSCustomObject]$config = fn_Read_JSON ".\config\gzen.json"
#$config | Format-List

$header = fn_Read_To_Array $config.orign_header

# *事前申請*.txt はSHIFT-JISでエンコードされている。
$pan_folder = (${HOME} + $config.temp_folder)
#$config.gZEN_targets
$values_filenames = fn_Find_Files $pan_folder $config.gZEN_targets
#$values_filenames

$values = foreach ($_ in $values_filenames) {
  fn_Read_To_Array $_
  # 使用済みの事前承認ファイルを waste フォルダへ移動する
  #Move-Item -Path $_ -Destination ($pan_folder + $config.waste_folder)
}

#テキストファイルを読み込み、ヘッダーをつけてCSVファイルを生成。
$csv_obj = fn_Bind_As_CSV $header $values

$sorted_field = fn_Read_To_Array $config.sorted_header

# 必要な項目の情報を抽出する
[PSCustomObject[]]$final_list = fn_Shape_Values $csv_obj $sorted_field $config.app_coms_path
$final_list

$CSV_destination = (${HOME} + $config.export_CSV_path)
$JSON_destination = (${HOME} + $config.export_JSON_path)
#$destination
foreach ($_ in @($CSV_destination, $JSON_destination)) {
  if (!(Test-Path $_)) {
    # 空ファイルを作る
    New-Item -Path $_ -ItemType File -Force
  }
}
# CSVファイルを出力
fn_Export_CSV $final_list $CSV_destination
# JSONファイルを出力
fn_Export_JSON $final_list $JSON_destination

#clipboardに出力する
<#
$private:comma = ', '
$private:tab = "`t"
$plain_text = Get-Content -Path $CSV_destination -Encoding UTF8
$formatted_text = $plain_text.Replace('"', '')
$formatted_text.Replace($comma, $tab) | Set-Clipboard
#>


#ftの親パスを取得
$private:app_path = Convert-Path .
$private:cpy_path = ($app_path + $config.cpy_path_monolith)
#Test-Path($cpy_path)
$private:command_name = $config.command_name

# EDGEブラウザ起動し、cpyを起動する。
Start-Process msedge $cpy_path

# 通知を表示
notifycation $command_name ("🐈., 💩💩, , .  💩, 🌲🏡  Ver " + $config.version)
