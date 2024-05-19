<#

ソースとなるcsvファイルの先頭行を削除し、
別名csvファイルに書き出す。
ソースとなったcsvファイルは別フォルダへ移動する。
#>

# Cut off Head. 最初の一行目を削除する。
Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

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

$source_path = (${HOME} + "\Downloads\PAN\登録者管理リスト.csv")
$text = Get-Content -Path $source_path -Encoding Default

$export_path = (${HOME} + "\Downloads\PAN\登録者管理リスト_coh.csv")
Set-Content -Path $export_path $text[2..($text.Length - 1)] -Encoding Default

$waste_folder = (${HOME} + "\Downloads\PAN\waste\csv")

$not_exisits_waste_folder = !(Test-Path $waste_folder)

# $source_path は不要なので移動する
if($not_exisits_waste_folder){
  Write-Host "\waste\csvフォルダを作る"
  New-Item -Path $waste_folder -ItemType Directory -Force
}
Move-Item -Path $source_path -Destination $waste_folder -Force

fn_notifycation ("CutOff Head...") ("完了 : 🐶 🐶 🐶 ")

