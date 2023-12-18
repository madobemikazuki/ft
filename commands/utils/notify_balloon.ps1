param(
  [String]$title = "コマンド名",
  [String]$message = "🐈.,💩💩,,.  💩,  🌲🏡"
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