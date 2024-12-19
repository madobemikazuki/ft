Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

# 低頻度でファイルの変更を感知したい場合に利用する。
# システムリソースの消費を小さくしなければならない環境で利用する。
# 大量のファイルを監視するにはリソースコストを考慮すること。
class FT_Ezy_Watcher {

  static Start([HashTable[]]$_orders) {
    Write-Host "ファイル監視を終了するときは CTRL + C を押下してください。"
    $lastCheck = @{}

    while ($True) {

      foreach ($_order in $_orders) {
        $_path = $_order.path
        $currentFiles = Get-ChildItem -Path $_path | Select-Object -Property FullName, LastWriteTime

        if ($lastCheck.ContainsKey($_path)) {
          foreach ($_file in $currentFiles) {
            $lastFile = $lastCheck[$_path] | Where-Object { $_.FullName -eq $_file.FullName }
            if ( (-not $lastFile) -or ($lastFile.LastWriteTime -ne $_file.LastWriteTime)) {
              #ここでスクリプトブロックを展開できる。
              Write-Host "changed:: $($_file.FullName)"
              & $_order.action_block
            }
          }
        }
        $lastCheck[$_path] = $currentFiles
      }
      Start-Sleep -Second 10
    }
  }


  <#
  static Start([String[]]$_paths, [ScriptBlock]$_action_block) {
    Write-Host "ファイル監視を終了するときは CTRL + C を押下してください。"
    $lastCheck = @{}
    while ($True) {
      foreach ($_path in $_paths) {
        $currentFiles = Get-ChildItem -Path $_path -Recurse | Select-Object -Property FullName, LastWriteTime

        if ($lastCheck.ContainsKey($_path)) {
          foreach ($_file in $currentFiles) {
            $lastFile = $lastCheck[$_path] | Where-Object { $_.FullName -eq $_file.FullName }
            if ( (-not $lastFile) -or ($lastFile.LastWriteTime -ne $_file.LastWriteTime)) {
              
              #ここでスクリプトブロックを展開できる。
              & $_action_block
              Write-Host "changed:: $($_file.FullName)"
            }
          }
        }
        $lastCheck[$_path] = $currentFiles
      }
      Start-Sleep -Second 10
    }
  }
#>
}

