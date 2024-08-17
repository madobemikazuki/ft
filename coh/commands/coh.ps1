<#

ソースとなるcsvファイルの先頭行を削除し、
別名csvファイルに書き出す。
ソースとなったcsvファイルは別フォルダへ移動する。
#>

# Cut off Head. 最初の一行目を削除する。
Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

function fn_Read {
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

function script:fn_notifycation {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$title,
    [Parameter(Mandatory = $True, Position = 1)]
    [String]$message
  )
  Add-Type -AssemblyName System.Windows.Forms
  $MUTEX_NAME = "Global\mutex" #多重起動チェック用

  try {
    $mutex = New-Object System.Threading.Mutex($False, $MUTEX_NAME)
    #多重起動チェック
    if ($mutex.WaitOne(0, $False)) {
      $notify_icon = New-Object Windows.Forms.NotifyIcon

      $notify_icon.Icon = [Drawing.SystemIcons]::Application
      $notify_icon.Visible = $True

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

function script:fn_Write_JSON {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$_path,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject]$_Object,
    [Parameter(Mandatory = $True, Position = 2)]
    [System.Text.Encoding]$_encoding
  )
  #既存するファイルを上書きする
  if (Test-Path $_path) {
    New-Item -Path $_path -ItemType File -Force
  }
  [System.IO.File]::WriteAllLines($_path, (ConvertTo-Json $_Object), $_encoding)
}

function fn_LookUpHash {
  Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [PSCustomObject[]]$_Object_List,
  [Parameter(Mandatory = $True, Position = 1)]
  [String]$_key
  )
  $LookUpHash = [ordered]@{}
  foreach($_ in $_Object_List){
    $LookUpHash[$_.$_key] = $_
  }
  return [PSCustomObject]$LookUpHash
}

$config = fn_Read ".\config\coh.json"

$source_path = (${HOME} + $config.source_path)
$text = Get-Content -Path $source_path -Encoding Default

$export_csv_path = (${HOME} + $config.export_csv_path)
# $text の3行目から $textの最終行まで取得して、作成したファイルに書き込む
$text_list = $text[2..($text.Length - 1)]
Set-Content -Path $export_csv_path $text_list -Encoding Default

[PSCustomObject[]]$obj_list = fn_Read $export_csv_path
$hash_key = $config.hash_key
$LookUpObject = fn_LookUpHash $obj_list $hash_key
$utf8_with_BOM = New-Object System.Text.UTF8Encoding $True
$export_json_path = (${HOME} + $config.export_json_path)
fn_Write_JSON $export_json_path $LookUpObject $utf8_with_BOM

$waste_folder = (${HOME} + $config.waste_folder)
$not_exisits_waste_folder = !(Test-Path $waste_folder)

# $source_path は不要なので移動する
if($not_exisits_waste_folder){
  Write-Host "\waste\csvフォルダを作る"
  New-Item -Path $waste_folder -ItemType Directory -Force
}
Move-Item -Path $source_path -Destination $waste_folder -Force

fn_notifycation ("CutOff Head...") ("完了 : 🐶 🐶 🐶 ")

